# This script recursively gets membership of a DL and then saves a text file of the membership.
# Use this script if someone needs to visually look at the hierarchy or membership and/or needs to see whether there are sender restrictions or allows external senders.

## Set Variables:
    $group = ""
    $members = New-Object System.Collections.ArrayList
    $indentLine = "- "
    $global:userCount = 0
    $global:groupCount = -1
    $includeUsers = $true

## Create the Function
function getMembership($group) {
Write-Host "Getting membership for $group..."
    
$searchGroup = Get-DistributionGroupMember ($group.ToString()) -ResultSize Unlimited | Sort-Object -property @{Expression="recipienttype";Descending=$true}, @{Expression="displayname";Descending=$false}

    foreach ($member in $searchGroup) {
        
        # It's a Group
        if ($member.RecipientTypeDetails -match "Group" -and $member.DisplayName -ne "") {
            $theCurrentGroup = Get-DistributionGroup $member.identity
            $duplicateGroup = $false
            $global:groupCount ++
            $restrictedSendersString = ""
            
            #Check if this group has already been seen
            if ($members -match $member.DisplayName) {
                $duplicateGroup = $true
            }
                  
            ####
            $groupNotes = ""
                
            if ($theCurrentGroup.RequireSenderAuthenticationEnabled -eq $false) {
                $groupNotes += " **"
            }
                  
            if ([string]$theCurrentGroup.AcceptMessagesOnlyFromSendersOrMembers -notlike $null) {
                $restrictedSenders = $theCurrentGroup.AcceptMessagesOnlyFromSendersOrMembers
                $restrictedSendersString = " (Senders restricted to: "
    
                $restrictedSendersString += toListofNames($restrictedSenders)
    
                $restrictedSendersString += ")"
                $groupNotes += $restrictedSendersString
            }  

            # add line to the array for this group
            $members.Add($indentLine+$member.DisplayName+$groupNotes) >$null
            
            $indentLine="-"+$indentLine

            if ($duplicateGroup -eq $true) {
                Write-Host "Found Duplicate Member Group:" $member.DisplayName
                if ($includeUsers -eq $true) {
                    $members.Add($indentLine+"Duplicate Group - Membership of this group has already been listed.") >$null
                    }
            }            
            #get the membership of this group
            else {
                getMembership($member.DisplayName)
            }                  
            $indentLine = $indentLine.substring(1,$indentLine.length-1)
        } 
                
        # It's a user                
        else {
            
            if ($member.DisplayName -ne "") {
                if ($includeUsers -eq $true) {
                    $members.Add($indentLine+$member.DisplayName) >$null
                    }
                $global:userCount ++
                }
        }
    }
    }


#Function to convert the Exchange MultiValue Property of user CNs to a string of names
function toListOfNames ($MultiProperty) {

    $NameString = ""
    foreach ($MemberProp in $MultiProperty) {
        #$userInfo = Get-User ($MemberProp.tostring())
        $NameString += (Get-Recipient $MemberProp).DisplayName + ";"
        #Write-Host $NameString
        }
    $NameString = $NameString.TrimEnd(";")
    return $NameString
    }

##################START

$group = Read-Host "Hi. What group would you like membership for? "
$includeUsersResponse = Read-Host "Should users be included in the report? (Y)ES or (N)o "

if ($includeUsersResponse -eq "N") {
    $includeUsers = $false
    }

$requestedGroup = Get-DistributionGroup ($group.ToString())


##### Add group name and notes for root group
$groupNotes = ""
if ($requestedGroup.RequireSenderAuthenticationEnabled -eq $false) {
    $groupNotes += " **"
    }
if ([string]$requestedGroup.AcceptMessagesOnlyFromSendersOrMembers -notlike $null) {
    $restrictedSenders = $requestedGroup.AcceptMessagesOnlyFromSendersOrMembers
    $restrictedSendersString = " (Senders restricted to: "
    
    $restrictedSendersString += toListofNames($restrictedSenders)
    
    $restrictedSendersString += ")"
    $groupNotes += $restrictedSendersString
    }
           
    # add line to the array for this group
    $members.Add($requestedGroup.DisplayName + $groupNotes) >$null

####

getMembership($requestedGroup)



## Output results
if ($global:groupCount -eq -1) {
    $global:groupCount ++
    }
$fileText = "Membership of " + $requestedGroup.DisplayName + " " + (get-date).tostring() + " : Users= " + $global:userCount + " Groups= " + $global:groupCount
$fileText = $fileText + "`r`n** indicates that the group allows unauthenticated senders. Messages from senders outside the organization will be allowed.`r`n"

$filePath = "C:\Users\$env:USERNAME\Documents\Membership of "+$group+".txt"
$fileText | Out-File $filePath

$members.GetEnumerator() | Out-File -Append $filePath

write-host $fileText
write-host "Written to " $filePath
