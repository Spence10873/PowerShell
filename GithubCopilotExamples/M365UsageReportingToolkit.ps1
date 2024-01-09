## This script is meant to provide blocks of code that can be used to interact with Microsoft 365 Usage Reporting API.
## The Usage Reporting API is a REST API that can be used to retrieve information about the usage of Microsoft 365 services.
## The Usage Reporting API is documented here: https://learn.microsoft.com/en-us/graph/api/resources/report?view=graph-rest-1.0

## Import the necessary modules to authenticate to the Microsoft Graph PowerShell SDK and to make calls to Usage Reporting resources in the Graph API
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Reports

## Authenticate to the Microsoft Graph PowerShell SDK
Connect-Graph -Scopes "Reports.Read.All"

## Get a period report of the count of daily active Office 365 users
$periodId = "D7"
Get-MgReportOffice365ActiveUserCount -Period $periodId

## Get a period report of the count of daily active Office 365 users for a specific date range
$periodId = "D7"
$startDate = "2021-01-01"
$endDate = "2021-01-31"
Get-MgReportOffice365ActiveUserCount -StartDate $startDate -EndDate $endDate

## Get a period report of the count of daily active Office 365 users for a specific date range and group by a specific property
$periodId = "D7"
$startDate = "2021-01-01"
$endDate = "2021-01-31"
$groupBy = "userPrincipalName"
Get-MgReportOffice365ActiveUserCount -StartDate $startDate -EndDate $endDate -GroupBy $groupBy
