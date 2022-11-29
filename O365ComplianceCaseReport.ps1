
<#
Import-Module C:\Scripts\Functions\Write-Log.ps1

[string]$LogFileDate = (get-date -f s).replace(":","-")
$LogFileName = "C:\Scripts\ScheduledTasks\O365ComplianceCases\Logs\O365ComplianceCaseReport_$LogFileDate.log"

Write-Log "Creating credential object for O365ExchangeAdmin"
$O365ExchangeAdminPW = "01000000d08c9ddf0115d1118c7a00c04fc297eb01000000dc85d1ab3700a54ab118e8d87c19c30e0000000002000000000003660000c000000010000000ac89cecb21607793c5f50e21898c57830000000004800000a0000000100000002ef6af28259fdcf2b909726be8d35a4418000000ce8b001a8672dbdef442551fce4d48b2d5b5671dd4c475361400000038b25a7170fe1d5a2a9e913c29052b02490ec7cc"
$O365ExchangeAdminSecure = $O365ExchangeAdminPW | convertto-secureSTring
$creds = New-Object System.Management.Automation.PSCredential ("O365ExchangeAdmin@cerner.net", $O365ExchangeAdminSecure)
Write-Log "Connecting to Security and Compliance Center"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://nam01b.ps.compliance.protection.outlook.com/PowerShell-LiveId -Credential $Creds -Authentication Basic -AllowRedirection
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Creds -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber -DisableNameChecking

Write-Log "Getting Compliance Cases"

#>
$ReportDate = (get-date -f s).replace(":","-")
$Cases = Get-ComplianceCase
Write-Host "There are currently $($Cases.count) compliance cases in Office 365"

$ActiveCases = $Cases | Where{$_.Status -eq "Active"}
Write-Host "Of those $($Cases.count) cases, $($ActiveCases.count) are active"

Write-Host "Building table to include in report"

$ActiveCasesTable = @()
$i = 0
$ActiveCasesCount = $ActiveCases.count
Foreach($ActiveCase in $ActiveCases)
    {
    $i++
    Write-Progress -Activity "Building a table of each active case, including CaseName, CreatedDateTime, LastModifiedDateTime, Status, ExchangeLocation, and SharePointLocation" -Status "Working on $($ActiveCase.Name), ($i / $ActiveCasesCount)" -PercentComplete (($i / $ActiveCasesCount) * 100)
    $ActiveCaseHoldPolicy = $null
    $ActiveCaseWorkLoad = $null
    $ExchangeLocation = $null
    $SharePointLocation = $null
    $ActiveCaseHoldPolicy = Get-CaseHoldPolicy $ActiveCase.name
    If($ActiveCaseHoldPolicy)
        {
        If($ActiveCaseHoldPolicy.Workload -like "*Exchange*")
            {
            $ExchangeLocation = $ActiveCaseHoldPolicy.ExchangeLocation.Name
            }
        If($ActiveCaseHoldPolicy.Workload -like "*SharePoint*")
            {
            $SharePointLocation = $ActiveCaseHoldPolicy.SharePointLocation.Name
            }
        }
    $TempTable = New-object System.Object
    $TempTable | add-member -type NoteProperty -Name CaseName -Value $ActiveCase.Name
    $TempTable | add-member -type NoteProperty -Name CreatedDateTime -Value $ActiveCase.CreatedDateTime
	$TempTable | add-member -type NoteProperty -Name LastModifiedDateTime -Value $ActiveCase.LastModifiedDateTime
    $TempTable | add-member -type NoteProperty -Name Status -Value $ActiveCase.Status
    $TempTable | add-member -type NoteProperty -Name ExchangeLocation -Value $ExchangeLocation
    $TempTable | add-member -type NoteProperty -Name SharePointLocation -Value $SharePointLocation

    $ActiveCasesTable += $TempTable

    }

$CSVName = "C:\Temp\O365ComplianceCaseReport_$ReportDate.csv"

Write-Host "Table has been built, and will now be exported to $CSVName"
$ActiveCasesTable | Sort-Object CreatedDateTime -Descending | Export-Csv -NoTypeInformation $CSVName

$MessageBody = @"
There are currently $($Cases.count) compliance cases in Office 365. Of those $($Cases.count) cases, $($ActiveCases.count) are active.
Attached is a CSV containing each active case.
"@

Write-Host "CSV has been created, now attaching it to an email and sending it"
Send-MailMessage -From "O365ComplianceCaseReporting@contoso.com" -To "ComplianceTeam@contoso.com", "ComplianceDirector@contoso.com" -Subject "Office 365 Compliance Cases Report" -Body $MessageBody -Attachments $CSVName -SmtpServer "smtp.contoso.com"
