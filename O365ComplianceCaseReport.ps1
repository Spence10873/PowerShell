
$ReportDate = (get-date -f s).replace(":","-")
$Cases = Get-ComplianceCase
Write-Host "There are currently $($Cases.count) compliance cases in Office 365"

$ActiveCases = $Cases | Where-Object{$_.Status -eq "Active"}
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
