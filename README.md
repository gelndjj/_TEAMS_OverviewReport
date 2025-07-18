# 📊 _TEAMS_OverviewReport

PowerShell script to generate a **fast, detailed overview** of all Microsoft Teams in your tenant using **batched Microsoft Graph API** calls. Ideal for audits, governance reviews, or admin reports.

---

## 📁 Output

The script generates a `.csv` file named like:
Teams_Overview_20250710_1933.csv


Each row represents a Team with metadata and settings extracted from the Microsoft Graph.

---

## 🔍 Output Columns

| Column                        | Description                                                                 |
|------------------------------|-----------------------------------------------------------------------------|
| `Id`                         | Unique Team ID (same as Group ID)                                           |
| `DisplayName`                | Display name of the Team                                                   |
| `Description`                | Description of the Team                                                    |
| `Classification`             | Classification label (if any)                                              |
| `Visibility`                 | Public or Private visibility                                               |
| `Specialization`             | Type of Team (e.g., educationClass, unknownFutureValue)                    |
| `WebUrl`                     | Web link to open the Team                                                  |
| `CreatedDateTime`            | When the Team was created                                                  |
| `IsArchived`                 | Whether the Team is archived                                               |
| `IsMembershipLimitedToOwners`| Whether only owners can see members                                        |
| `ShowInSearch`               | Visibility in search and suggestions                                       |
| `OwnersCount`               | Number of Team owners                                                     |
| `MembersCount`              | Number of internal members                                                |
| `GuestsCount`               | Number of guest members                                                   |
| `OwnerUPNs`                 | List of owners’ UPNs, semicolon-separated                                 |
| `ChannelCount`              | Count of standard channels                                                |
| `PrivateChannelCount`      | Count of private channels                                                 |
| `SharedChannelCount`       | Count of shared channels                                                  |
| `AllowCreateUpdateChannels`| Member permission: create/update channels                                 |
| `AllowCreatePrivateChannels`| Member permission: create private channels                                |
| `AllowDeleteChannels`      | Member permission: delete channels                                        |
| `AllowAddRemoveApps`       | Member permission: manage apps                                            |
| `AllowCreateUpdateRemoveTabs`| Member permission: manage tabs                                           |
| `AllowCreateUpdateRemoveConnectors`| Member permission: manage connectors                           |
| `AllowUserEditMessages`    | Messaging permission: edit own messages                                   |
| `AllowUserDeleteMessages`  | Messaging permission: delete own messages                                 |
| `AllowOwnerDeleteMessages` | Messaging permission: owners delete messages                              |
| `AllowTeamMentions`        | Allow @Team mentions                                                      |
| `AllowChannelMentions`     | Allow @Channel mentions                                                   |
| `AllowGiphy`               | Allow Giphy in chats                                                      |
| `GiphyContentRating`       | Giphy rating allowed (Moderate, Strict)                                   |
| `AllowStickersAndMemes`    | Allow stickers and memes                                                  |
| `AllowCustomMemes`         | Allow users to add their own memes                                        |

---

## ⚙️ Requirements

- PowerShell 7+
- Microsoft Graph PowerShell SDK  
- **Permissions required** (delegated or managed identity):

Group.Read.All
Team.ReadBasic.All
Channel.ReadBasic.All
ChannelMember.Read.All
TeamMember.Read.All
User.Read.All


> ✅ This version uses batching to reduce time and API throttling risks.

---

## 🚀 How to Run

1. Open PowerShell
2. Run the script

```pwsh
.\Teams_Overview_Report.ps1
```
CSV will be generated in the same folder.

---

## ⏱️ Runtime

Execution time is displayed at the end (in seconds) to track performance across tenants.

---

## 📤 Sample Use Cases

Audit and governance reporting
Teams cleanup and optimization
Security review of private/shared channels
Admin ownership and team visibility checks

---

## 🔐 Security Note

No tenant-specific hardcoding.
Authentication uses interactive sign-in via Connect-MgGraph.

---

## ☁️ Azure Automation Runbook

This report can be scheduled to run automatically using Azure Automation with a Managed Identity and have the resulting CSV uploaded to SharePoint.


## 🔧 Setup Instructions

1. Required PowerShell Modules in Azure Automation:

Microsoft.Graph
PnP.PowerShell

2. Required Microsoft Graph Permissions (Application):

Group.Read.All
Team.ReadBasic.All
Channel.ReadBasic.All
ChannelMember.Read.All
TeamMember.Read.All
User.Read.All

3. SharePoint Access:
Assign the Managed Identity as a Contributor on the SharePoint document library where the report will be saved.

---

## 🛠️ Runbook Features

Authenticates using Managed Identity (no credentials needed)
Retrieves metadata and settings for all Microsoft Teams
Uses batched Microsoft Graph requests to minimize throttling
Generates a detailed .csv file:
```plain
Teams_Overview_Report.csv
```
Uploads the file to a specified SharePoint document library

---

## ✏️ Customization
Update the following lines in the runbook script to reflect your SharePoint environment:

```pwsh
$sharePointSiteUrl = "https://<YourTenant>.sharepoint.com/sites/<YourSiteName>"
$sharePointLibraryPath = "Shared Documents/Reporting/Teams"
```
