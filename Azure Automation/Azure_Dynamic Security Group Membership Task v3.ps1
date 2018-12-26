<#
.Synopsis
   Adds users/computer to groups dynamically
.DESCRIPTION
   Dynamic process for security group memberships on conference rooms, executives, DS, etc
.NOTES
   Running on Azure Automation runbook
#>



$DateFormatted = Get-Date -Format MMM.dd.yyyy-hhmm
$Path = "$($DateFormatted).log"
#$Path = "C:\users\christopher.reese\desktop\Dynamic Security Group Log $($DateFormatted).log"
$LogFile = New-Item -Path $Path
$FinalResults = @()

#####################################################################################
#Functions
#####################################################################################
#Add workstation to group
Function Add-ComputerToGroup ($Computer_NAME, $SecurityGroup) {
    Add-ADPrincipalGroupMembership -Identity ($Computer_NAME + '$') -MemberOf "$SecurityGroup" -ErrorAction SilentlyContinue #-WhatIf
}
#Add user to group
Function Add-UserToGroup ($Username, $SecurityGroup) {
    Add-ADPrincipalGroupMembership -Identity $Username -MemberOf "$SecurityGroup" -ErrorAction SilentlyContinue #-WhatIf
}
#Get direct reports recursively
Function Get-ADDirectReports ($samaccountname) {
    Get-ADUser $samaccountname -Properties directreports | % {
        $_.directreports | foreach -Process {
            Get-ADUser $PSItem -Properties manager | select name,samaccountname,@{l="manager";e={ (Get-ADUser $PSItem.manager).samaccountname } }
            $results = Get-ADDirectReports $PSItem
        }
    }
    Return $results
}


#####################################################################################
#Add conference rooms to screen saver exclusions list
#####################################################################################
"##  Add conference rooms to screen saver exclusions list  ##" | Out-File -FilePath $Path -Append

$ConferenceWorkstations = @()
$ConferenceRoomOUs = Get-ADOrganizationalUnit -filter 'name -like "*Conf*"' 
$ScreenSaverExclusionGroup = "Screen Saver Exclusion List"
$ScreenSaverExclusionGroupMembers = Get-ADGroupMember $ScreenSaverExclusionGroup

#Get all workstations
foreach ($OU1 in $ConferenceRoomOUs) {
     foreach ($PC1 in (Get-ADComputer -filter * -SearchBase $OU1)) {
     $ConferenceWorkstations += $pc1.Name
     }
}

#Add workstations to group
#$conferenceworkstations | % {Add-ComputerToGroup $_ $ScreenSaverExclusionGroup}

#Add workstations to group
foreach ($ConfWorkstation in $ConferenceWorkstations) {
    if ((![string]::IsNullOrWhiteSpace($ConfWorkstation)) -and ($ConfWorkstation | ? {$ScreenSaverExclusionGroupMembers.name -notcontains $_})) {
        write-output "Adding...[$($ConfWorkstation)] to [$($ScreenSaverExclusionGroup)]"
        Add-ComputerToGroup $ConfWorkstation $ScreenSaverExclusionGroup
        "Adding $($ConfWorkstation) to group: $($ScreenSaverExclusionGroup)" | Out-File -FilePath $Path -Append
        $Obj = [pscustomobject] @{
            Computername = $ConfWorkstation
            Groupname = $ScreenSaverExclusionGroup
        }
        $FinalResults += $Obj
    }
}



#####################################################################################
#Add executives and local admin exclusion groups
#####################################################################################
"##  Add executives and local admin exclusion groups  ##" | Out-File -FilePath $Path -Append
$ExecutiveWorkstations = @()
$ExecutiveOUs = Get-ADOrganizationalUnit -filter 'name -like "*Executiv*"' -SearchBase ""
$LocalAdminExclusionList = "GPO - Local Admin Fix - Exempt"
$LocalAdminExclusionGroupMembers = Get-ADGroupMember $LocalAdminExclusionList

#Get all executive computers
foreach ($OU2 in $ExecutiveOUs) {
     foreach ($PC2 in (Get-ADComputer -filter * -SearchBase $OU2)) {
     $ExecutiveWorkstations += $pc2.Name
     }
}

#Add workstations to group
foreach ($ExecWorkstation in $ExecutiveWorkstations) {
    if ((![string]::IsNullOrWhiteSpace($ExecWorkstation)) -and ($ExecWorkstation | ? {$LocalAdminExclusionGroupMembers.name -notcontains $_})) {
        write-output "Adding...[$($ExecWorkstation)] to [$($LocalAdminExclusionList)]" 
        Add-ComputerToGroup $ExecWorkstation $LocalAdminExclusionList
        "Adding $($ExecWorkstation) to group: $($ExecutiveWorkstations)" | Out-File -FilePath $Path -Append
        $Obj = [pscustomobject] @{
            Computername = $ExecWorkstation
            Groupname = $LocalAdminExclusionList
        }
        $FinalResults += $Obj
    }
}



#####################################################################################
#Add digital support to admin security group
#####################################################################################
"##  Add digital support to admin security group  ##" | Out-File -FilePath $Path -Append

$DigitalSupportUsers = Get-ADDirectReports "<username>"
$DigitalSupportAdminGroup = "Digital.Support.Admins"
$DigitalSupportAdmins = Get-ADGroupMember $DigitalSupportAdminGroup

foreach ($DPUser in $DigitalSupportUsers) {
    if ((![string]::IsNullOrWhiteSpace($DPuser)) -and ($DPuser | ? {$DigitalSupportAdmins.samaccountname -notcontains $_.samaccountname})) {
        #write-output "Adding...[$($DPUser.samaccountname)] to [$($DigitalSupportAdminGroup)]" 
        Add-UserToGroup $DPuser.samaccountname $DigitalSupportAdminGroup
        "Adding $($DPUser.samaccountname) to group: $($DigitalSupportAdminGroup)" | Out-File -FilePath $Path -Append
        Write-Output "Adding $($DPUser.samaccountname) to group: $($DigitalSupportAdminGroup)"
        $Obj = [pscustomobject] @{
            Computername = $DPUser.samaccountname
            Groupname = $DigitalSupportUsers
        }
        $FinalResults += $Obj
    }
}


#####################################################################################
#Add engineering group to admin security group
#####################################################################################
"##  Add engineering group to admin security group  ##" | Out-File -FilePath $Path -Append

$EngineerSupportUsers_Combined = @()
$EngineerSupportUsers1 = Get-ADDirectReports "<username>"
$EngineerSupportUsers2 = Get-ADDirectReports "<username>"
$EngineerSupportUsers_Combined += $EngineerSupportUsers1
$EngineerSupportUsers_Combined += $EngineerSupportUsers2
$EngineerSupportAdminGroup = "Engineer.Support.Admins"
$EngineerSupportAdmins = Get-ADGroupMember $EngineerSupportAdminGroup

#Filter out blanks, duplicates, Cory's team
foreach ($EngUser in $EngineerSupportUsers_Combined) {
    if (
        (![string]::IsNullOrWhiteSpace($Enguser)) -and 
        ($Enguser | ? {$EngineerSupportAdmins.samaccountname -notcontains $_.samaccountname}) -and 
        ($EngUser.samaccountname -ne "cory.kolb") -and
        ($EngUser.manager -ne "cory.kolb") -and 
        ($EngUser.manager -ne "dmitry.markin")) {
            write-output "Adding...[$($EngUser.samaccountname)] to [$($EngineerSupportAdminGroup)]" 
            Add-UserToGroup $EngUser.samaccountname $EngineerSupportAdminGroup
            "Adding $($EngUser.samaccountname) to group: $($EngineerSupportAdminGroup)" | Out-File -FilePath $Path -Append
            $Obj = [pscustomobject] @{
                Computername = $EngUser.samaccountname
                Groupname = $EngineerSupportUsers_Combined
            }
            $FinalResults += $Obj
    }
}


#####################################################################################
#Compile and email results
#####################################################################################

if (![string]::IsNullOrEmpty($Results)) {
    #Send report via email
    Write-Output "Sending email..."
    $to = ""
    $from = $to
    $subject = "Security Group Assignments"
    $smtpserver = ""
    $body = "
    
    
    $($FinalResults | out-string)   
    
    "
    send-mailmessage -to $to -from $from -Subject $subject -body $body  -smtpserver $smtpserver 
}
else {
    "Not Sending email..." | Out-File -FilePath $Path -Append
    Write-Output "Not Sending email..."
}