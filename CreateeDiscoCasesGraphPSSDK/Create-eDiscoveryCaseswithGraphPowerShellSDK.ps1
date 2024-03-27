Function Get-RandomUPN {
    #A simple array of UPNs is needed, no matter how you get them
    $UPNList = Get-Content c:\temp\ehlofromkc_emails.txt
    $UPNNumber = Get-Random -Minimum 0 -Maximum $UPNList.count
    return $UPNList[$UPNNumber]
}

#Specify the number of cases and custodians you want to create
$CaseCount = 200
$MinimumCustodianCount = 4
$MaximumCustodianCount = 12

#Iterate through creation of each new case
For($i = 0; $i -lt $CaseCount; $i++) {

    Write-Progress -Activity "Creating $CaseCount cases, each with $MinimumCustodianCount to  $MaximumCustodianCount custodians" -Status "Working on ($i / $CaseCount)" -PercentComplete (($i / $CaseCount) * 100)
    #Create Case
    $NewCase = New-MgComplianceEdiscoveryCase -DisplayName "MGCase$i"
    $CaseID = $NewCase.Id
    
    #Randomly determine how many custodians the case will have
    $CustodianCount = Get-Random -Minimum $MinimumCustodianCount -Maximum $MaximumCustodianCount

    #Generate list of random custodian UPNs
    $CustodianList = @()
    Do{
        $CustodianList += Get-RandomUPN
    } Until (($CustodianList | Get-Unique).count -eq $CustodianCount)

    #Add custodians to case
    Foreach($CustodianUPN in $CustodianList) {
        $Custodian = New-MgComplianceEdiscoveryCaseCustodian -CaseId $CaseID -Email $CustodianUPN -ApplyHoldToSources:$True
        New-MgComplianceEdiscoveryCaseCustodianUserSource  -CaseId $CaseID -CustodianId $Custodian.ID -Email $CustodianUPN -IncludedSources "mailbox,site"
    }

    #Place all custodians of the case on hold 
    Add-MgComplianceEdiscoveryCaseCustodianHold -CaseId $CaseID
}