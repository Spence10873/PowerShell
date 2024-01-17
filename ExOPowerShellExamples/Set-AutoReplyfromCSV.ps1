## This script is a sample meant to show how to input a CSV and use it to set the auto reply for a list of users.
## It is not meant to be used in production as is.
## It is meant to be used as a starting point for your own scripts.
## It is not supported by Microsoft.
## It is provided AS IS without warranty of any kind.

## This script requires the Exchange Online PowerShell V3 module.

## Connect to Exchange Online PowerShell
Connect-ExchangeOnline -UserPrincipalName "admin@ehlofromkc.com"

## Import the CSV file
$CSV = Import-Csv -Path "c:\temp\AutoReply.csv"

## Loop through the CSV file, setting the auto reply for each user that doesn't currently have one set
ForEach ($User in $CSV) {
    ## Add a progress bar
    Write-Progress -Activity "Setting Auto Reply for $User.DisplayName" -Status "Setting Auto Reply for $User.DisplayName" -PercentComplete (($CSV.IndexOf($User) + 1) / $CSV.Count * 100)

    $UserUPN = $User.UserPrincipalName
    $UserAutoReply = Get-MailboxAutoReplyConfiguration -Identity $UserUPN
    If ($UserAutoReply.AutoReplyState -eq "Disabled") {
        Set-MailboxAutoReplyConfiguration -Identity $UserUPN -AutoReplyState Enabled -ExternalMessage $User.ExternalMessage -InternalMessage $User.InternalMessage
    }
}

