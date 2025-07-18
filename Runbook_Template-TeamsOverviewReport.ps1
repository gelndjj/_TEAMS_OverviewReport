<#
.SYNOPSIS
    Runbook to generate a Teams Overview Report using batched Microsoft Graph queries and upload it to SharePoint.

.REQUIREMENTS
    - Modules:
        • Microsoft.Graph
        • PnP.PowerShell
    - Managed Identity Permissions:
        • Graph (Application):
            - Group.Read.All
            - Team.ReadBasic.All
            - Channel.ReadBasic.All
            - ChannelMember.Read.All
            - TeamMember.Read.All
            - User.Read.All
        • SharePoint Contributor access to target library
#>

# STEP 1 – Connect to Graph & SharePoint
Connect-MgGraph -Identity
Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/me" | Out-Null

$sharePointSiteUrl = "https://<YourTenant>.sharepoint.com/sites/<YourSiteName>"
$sharePointLibraryPath = "Shared Documents/Reporting/Teams"
Connect-PnPOnline -Url $sharePointSiteUrl -ManagedIdentity

# STEP 2 – Initialize
$GraphBase   = "https://graph.microsoft.com/v1.0"
$BatchSize   = 20
$PauseMS     = 300
$start       = Get-Date

# STEP 3 – Get all Teams
$teams = Get-MgTeam -All -Property "id,displayName"
Write-Output "→ Found $($teams.Count) Teams. Building batch requests..."

# STEP 4 – Build all requests
$reqs = foreach ($t in $teams) {
    @(
        @{ id = "team_$($t.Id)";     method = "GET"; url = "/teams/$($t.Id)" }
        @{ id = "channels_$($t.Id)"; method = "GET"; url = "/teams/$($t.Id)/channels" }
        @{ id = "owners_$($t.Id)";   method = "GET"; url = "/teams/$($t.Id)/members?\$filter=roles/any(r:r eq 'owner')" }
    )
}

# STEP 5 – Send batched requests
$responses = @{}
for ($i=0; $i -lt $reqs.Count; $i += $BatchSize) {
    $batch = $reqs[$i..([math]::Min($i+$BatchSize-1,$reqs.Count-1))]
    $body  = @{ requests = $batch } | ConvertTo-Json -Depth 8
    $resp  = Invoke-MgGraphRequest -Method POST -Uri "$GraphBase/`$batch" -Body $body -ContentType 'application/json'
    foreach ($item in $resp.responses) { $responses[$item.id] = $item.body }
    Start-Sleep -Milliseconds $PauseMS
}

# STEP 6 – Assemble report
$report = foreach ($t in $teams) {
    $teamID = $t.Id
    $team     = $responses["team_$teamID"]
    $channels = $responses["channels_$teamID"].value
    $owners   = $responses["owners_$teamID"].value

    $std  = ($channels | Where-Object membershipType -eq 'standard').Count
    $priv = ($channels | Where-Object membershipType -eq 'private' ).Count
    $shrd = ($channels | Where-Object membershipType -eq 'shared'  ).Count

    $ownerUPNs = $owners | ForEach-Object {
        try { (Get-MgUser -UserId $_.userId -Property userPrincipalName).userPrincipalName }
        catch { $_.displayName }
    }

    $sum = $team.summary
    [pscustomobject]@{
        Id                              = $team.id
        DisplayName                     = $team.displayName
        Description                     = $team.description
        Classification                  = $team.classification
        Visibility                      = $team.visibility
        Specialization                  = $team.specialization
        WebUrl                          = $team.webUrl
        CreatedDateTime                 = $team.createdDateTime
        IsArchived                      = $team.isArchived
        IsMembershipLimitedToOwners     = $team.isMembershipLimitedToOwners
        ShowInSearch                    = $team.discoverySettings.showInTeamsSearchAndSuggestions
        OwnersCount                     = $sum.ownersCount
        MembersCount                    = $sum.membersCount
        GuestsCount                     = $sum.guestsCount
        OwnerUPNs                       = $ownerUPNs -join ' ; '
        ChannelCount                    = $std
        PrivateChannelCount             = $priv
        SharedChannelCount              = $shrd
        AllowCreateUpdateChannels        = $team.memberSettings.allowCreateUpdateChannels
        AllowCreatePrivateChannels       = $team.memberSettings.allowCreatePrivateChannels
        AllowDeleteChannels              = $team.memberSettings.allowDeleteChannels
        AllowAddRemoveApps               = $team.memberSettings.allowAddRemoveApps
        AllowCreateUpdateRemoveTabs      = $team.memberSettings.allowCreateUpdateRemoveTabs
        AllowCreateUpdateRemoveConnectors= $team.memberSettings.allowCreateUpdateRemoveConnectors
        AllowUserEditMessages            = $team.messagingSettings.allowUserEditMessages
        AllowUserDeleteMessages          = $team.messagingSettings.allowUserDeleteMessages
        AllowOwnerDeleteMessages         = $team.messagingSettings.allowOwnerDeleteMessages
        AllowTeamMentions                = $team.messagingSettings.allowTeamMentions
        AllowChannelMentions             = $team.messagingSettings.allowChannelMentions
        AllowGiphy                       = $team.funSettings.allowGiphy
        GiphyContentRating               = $team.funSettings.giphyContentRating
        AllowStickersAndMemes            = $team.funSettings.allowStickersAndMemes
        AllowCustomMemes                 = $team.funSettings.allowCustomMemes
    }
}

# STEP 7 – Save to CSV and upload to SharePoint
$csvName = "Teams_Overview_Report.csv"
$tempPath = Join-Path -Path $env:TEMP -ChildPath $csvName
$report | Export-Csv -Path $tempPath -Delimiter ';' -NoTypeInformation -Encoding UTF8

Add-PnPFile -Path $tempPath -Folder $sharePointLibraryPath -NewFileName $csvName -Values @{}

$duration = (Get-Date) - $start
Write-Output "Report uploaded to SharePoint: $sharePointLibraryPath/$csvName"
Write-Output "Duration: $($duration.TotalSeconds) seconds"