## This script is meant to provide blocks of code that can be used to interact with Microsoft Graph API WITHOUT using the Microsoft Graph PowerShell SDK.
## The Graph API is a REST API that can be used to retrieve information about the Microsoft 365 ecosystem.
## The Graph API is documented here: https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0
## Instead of using the Graph SDK, this script uses Invoke-RestMethod to make calls to the Graph API.
## To handle authentication, this script uses Invoke-RestMethod to request an access token from the Microsoft Identity Platform (https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow).

## Create a function to get an access token from the Microsoft Identity Platform
function Get-AccessToken {
    # Define parameteres for the function
    param(
        [Parameter(Mandatory = $true)]
        [string]$tenantId,
        [Parameter(Mandatory = $true)]
        [string]$clientId,
        [Parameter(Mandatory = $true)]
        [string]$clientSecret
    )
    
    ## Define the parameters for the request
    $params = @{
        Uri         = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        Method      = "POST"
        ContentType = "application/x-www-form-urlencoded"
        Body        = @{
            client_id     = $clientId
            scope         = "https://graph.microsoft.com/.default"
            client_secret = $clientSecret
            grant_type    = "client_credentials"
        }
    }

    ## Make the request
    $response = Invoke-RestMethod @params

    ## Return the access token
    return $response.access_token
}

## Prompt the user for a credential, expecting the client ID to be in the username field and the client secret to be in the password field
#$appCredential = Get-Credential
## First check if $appCredential exists, and only prompt for credentials if it doesn't
if ($appCredential -eq $null) {
    $appCredential = Get-Credential
}

## Prompt the user for a tenant ID
$tenantId = Read-Host "Enter the tenant ID"

## Get an access token from the Microsoft Identity Platform
$accessToken = Get-AccessToken -tenantId $tenantId -clientId $appCredential.UserName -clientSecret $appCredential.GetNetworkCredential().Password

## Test the access token by making a call to the Graph API
$graphResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users" -Method Get -Headers @{Authorization = "Bearer $accessToken"}
$first100Users = $graphResponse.value

## Display the first 100 users
$first100Users | Out-GridView