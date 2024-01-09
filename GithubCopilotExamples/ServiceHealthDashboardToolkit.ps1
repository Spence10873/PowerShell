## This script is meant to provide blocks of code that can be used to interact with Microsoft's Service Health Dashboard (SHD) API.
## The SHD API is a REST API that can be used to retrieve information about the health of Microsoft's cloud services.
## The SHD API is documented here: https://learn.microsoft.com/en-us/graph/api/resources/servicehealth?view=graph-rest-1.0


## Import the necessary modules to authenticate to the Microsoft Graph PowerShell SDK and to make calls to Service Health resources in the Graph API
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Devices.ServiceAnnouncement

## Authenticate to the Microsoft Graph PowerShell SDK
Connect-Graph -Scopes "ServiceHealth.Read.All"

## Get the current status of all services
Get-MgServiceAnnouncementHealthOverview

## Get the current status of Exchange Online
Get-MgServiceAnnouncementHealthOverview -ServiceHealthId "Exchange Online"

## Get the current status of SharePoint Online
Get-MgServiceAnnouncementHealthOverview -ServiceHealthId "SharePoint Online"

## List all issues
Get-MgServiceAnnouncementIssue -All | Out-GridView