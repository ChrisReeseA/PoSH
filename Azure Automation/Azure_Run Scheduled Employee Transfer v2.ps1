<#
 # <#
   .Synopsis
      Run scheduled transfer process
   .DESCRIPTION
      Search for named variable from WebJEA, run on schedule (Sunday night)
   .OUTPUTS
      Outputs located in Automation Account in Azure
   .NOTES
      Delete tool available in case process needs cancelled.  Located in WebJEA
   .FUNCTIONALITY
      Approved transfer process
   #>

#Credentials for O365
####################################################################
$PW_Path = ""
$Key_Path = ""

$key = Get-Content -Path $Key_Path
$o365userName = "<username>"
$passwordText = Get-Content -Path $PW_Path
$securedPW = $passwordText | ConvertTo-SecureString -Key $Key  
$o365cred = New-Object pscredential -ArgumentList $o365userName, $securedPW
####################################################################


#Functions
#####################################################################################
function Convert-ToOUFormat ($UserDN) {
    $ConversionResultOU = ($UserDN -split ',' | select -skip 1) -join ','
    return $ConversionResultOU
}
function Convert-ManagerFormat ($UserMgr) {
    $ConversionResultMgr = ($UserMgr -split ',' | select -first 1).replace('CN=','')
    return $ConversionResultMgr
}
function AzureAD ($O365cred) {
        try {
            Connect-AzureAD -Credential $O365cred
        }
        catch {
            $ErrorMessage = $_.Exception.Message
        }
        return $ErrorMessage    
}
function Remove-AzureGroups ($UPN) {    
    #Parse searchstring
    if ($UPN -like "*@*") {
        $SearchString = $UPN -split '@' | select -first 1
    }
    else {
        $SearchString = $UPN
    }

    #Azure AD Account
    $AzureADAccount = Get-AzureADUser -SearchString $SearchString 

    #Fetch Azure groups
    $AzureGroups = $AzureADAccount | Get-AzureADUserMembership | ? {[string]::IsNullOrEmpty($_.onpremisessecurityidentifier)}

    #Remove membership from user account
    foreach ($r in $AzureGroups) {
        #Write-Host $r.objectid
        #Write-Host $AzureADAccount.objectid
        try {
            Write-Output "[  Removing  ] from Azure group [  $($r.displayname)  ]"
            Remove-AzureADGroupMember -ObjectId $r.objectid -MemberId $AzureADAccount.ObjectId

        }
        catch {
            Write-Error "###ERROR | Azure Security Groups### $($failitem): $($error)"
            $script:ErrorCount++
        }
    }
}
function Compare-SecurityGroups ($UserAccounts) {
    $Compare = @()

    foreach ($user in $UserAccounts) {
        try {
            $Item = Get-ADUser -filter {description -eq $user} -Properties memberof
            $Compare += $Item
            #Write-Host $item.samaccountname -ForegroundColor Cyan
        }
        catch {
            Write-Output "Cannot find $($user) in AD" -ForegroundColor Red
        }
    }
        
    $Result = $Compare[0] | Select -ExpandProperty MemberOf
    ForEach ($Index in (1..($Compare.Count - 1))) {
        $Diff = $Compare[$Index] | Select -ExpandProperty MemberOf
        
        If ($Result -and $Diff) {   
            $Result = Compare-Object -ReferenceObject $Result -DifferenceObject $Diff -IncludeEqual | Where SideIndicator -EQ "==" | Select -ExpandProperty InputObject
        }
        Else {   
            Write-Warning "No common groups found!"
        }
    }
    return $Result
}
#####################################################################################


#Declare variables
#####################################################################################
$ErrorCount = 0
#####################################################################################


#Get information from Azure
#####################################################################################
#Get scheduled variable from job
Write-Output "Logging in Azure..."
$LocalTranscript += "Logging in Azure..."
Login-AzureRmAccount -Credential $o365cred | Out-Null
AzureAD $o365cred

#Get working transfer variable
$GetAzureVariable = @{
    ResourceGroupName = ""
    AutomationAccountName = ""
}
#$AzureTransferVariable = (Get-AzureRmAutomationVariable @GetAzureVariable | ? {$_.name -like "Transfer-*"}).value
$AzureTransferVariable = Get-AzureRmAutomationVariable @GetAzureVariable | ? {$_.name -like "Transfer-*"}


#Process job
#####################################################################################
if (![string]::IsNullOrEmpty($AzureTransferVariable) -and ($azuretransfervariable.count -le 15)) {
    foreach ($ScheduledJob in $AzureTransferVariable) {
        #Declare variables
        $JobID = @()
        $CopyEmployeeIDs = @()
        $LocalTranscript = @()
        $NewOUPath = @()
        $CopyEmployees = @()
        $ScheduledJobValue = $ScheduledJob.value
        $ScheduledJobName = $ScheduledJob.Name
        $JobID = $ScheduledJobValue.TechID
        $CopyEmployeeIDs = $ScheduledJobValue.CopyAccess
    
        #Get AD information
        $ADAccount = Get-ADUser -Filter 'description -eq $JobID' -Properties title,distinguishedname,memberof,manager,description,department,office
        #Kill process if more than 1
        if ($ADAccount.count -gt 1) {exit}
        $ADAccountMgr = Convert-ManagerFormat $ADAccount.Manager
        $ADAccountOU = Convert-ToOUFormat $ADAccount.distinguishedname
        $ADAccountTitle = $ADAccount.Title
        
        #Fetch employee(s) from textbox
        foreach ($i in $CopyEmployeeIDs) {
            $CopyEmployees += (Get-ADUser -Filter 'description -eq $i' -Properties memberof,manager,department,title)
        }
        
        #Determine if single user or multiple, get security groups
        if (@($CopyEmployeeIDs).count -gt 1) {
            $CopyEmployees | % {
                #Write-Output $_.samaccountname
                #$LocalTranscript += $_.samaccountname
                $NewOUPath += Convert-ToOUFormat $_.distinguishedname
            }
        }
        else {
            $NewOUPath = Convert-ToOUFormat $CopyEmployees.distinguishedname
        }
        
        #Remove any duplicate OUs
        $NewOUPath = $NewOUPath | select -Unique | sort
        
        #Process user
        #####################################################################################
        #Remove all security groups from membership
        Write-Host "Removing AD group memberships..."
        $LocalTranscript += "Removing AD group memberships..."
        try {
            $ADAccount | % {
                $Samaccountname = $_.samaccountname
                $_.memberof | % {
                    Write-Output "[  Removing  ] $($SamAccountName) from security group [  $($_)  ]"
                    $LocalTranscript += "[  Removing  ] $($SamAccountName) from security group [  $($_)  ]"
                }
                $_.memberof | Remove-ADGroupMember -Members $_.distinguishedname -Confirm:$false -ErrorAction Stop #-WhatIf
            }
        }
        catch {
            $ErrorCount++
            Write-Error "###ERROR | Security Groups### $($failitem): $($error)"          
            $LocalTranscript += "###ERROR | Security Groups### $($failitem): $($error)"
        }
        
        #Remove o365 groups (teams, distros, etc)
        Write-Output "Removing O365 group memberships..."
        $LocalTranscript += "Removing O365 group memberships..."
        $O365UPN = $ADAccount.UserPrincipalName.Replace("","")
        Remove-AzureGroups $O365UPN
        
        #Update security principals
        Write-Output "Adding group memberships..."
        $LocalTranscript +="Adding group memberships..."
        Compare-SecurityGroups $CopyEmployeeIDs | % {
            Write-Output "[  Adding  ] $($ADAccount.SamAccountName) to group [  $($_)  ]"
            $LocalTranscript += "[  Adding  ] $($ADAccount.SamAccountName) to group [  $($_)  ]"
            $ADGroupObject = Get-ADGroup $_
            try {
                Add-ADPrincipalGroupMembership $ADAccount.SamAccountName -MemberOf $ADGroupObject.Name -ErrorAction Stop #-WhatIf
            }
            catch {
                #$error = $_.Exception.Message
                #$failitem = $_.Exception.ItemName
                $ErrorCount++
                Write-Error "###ERROR | Security Groups### $($failitem): $($error)"
                $LocalTranscript += "###ERROR | Security Groups### $($failitem): $($error)"
            }
        }
        
        #Move to proper OU
        #Check if one OU was listed for automation move, if not, generate error
        Write-Host "Moving to new Organizational Unit..."
        $LocalTranscript += "Moving to new Organizational Unit..."
        if ($NewOUPath.Count -eq 1) {
            Write-Output "[  Moving  ] $($ADAccount.SamAccountName) to new OU [  $($NewOUPath)  ]"
            $LocalTranscript += "[  Moving  ] $($ADAccount.SamAccountName) to new OU [  $($NewOUPath)  ]"
            $ADObj = Get-ADObject $NewOUPath
            try {
                $ADAccount | Move-ADObject -TargetPath $ADObj -ErrorAction Stop #-WhatIf
            }
            catch {
                $ErrorCount++
                Write-Error "###ERROR | OU### $($failitem): $($error)"
                $LocalTranscript += "###ERROR | OU### $($failitem): $($error)"
            }
        }
        else {
            Write-Error "##ERROR##  :  User's OU was not changed because there are multiple values for destination OU.  Additional action is required to complete this process"
            $LocalTranscript += "##ERROR##  :  User's OU was not changed because there are multiple values for destination OU.  Additional action is required to complete this process"
        }

        #Clean up and process email notification for ITSupport, grab job status
        #####################################################################################

        #Remove Azure variable after processing
        Write-Output "Removing variable $($ScheduledJob.Name)"
        $LocalTranscript += "Removing variable $($ScheduledJob.Name)"
        Remove-AzureRmAutomationVariable @GetAzureVariable -Name $ScheduledJob.Name #-WhatIf
        
        #Get Job Output for email transcription
        #Write-Output "Getting job output..."
        #$JobOutput = (Get-AzureRmAutomationJobOutput -ResourceGroupName AzureAutomationGroup -AutomationAccountName GLAutomationAccount -Id $JobID -Stream Any | Get-AzureRmAutomationJobOutputRecord).value
        #Write-Output "Output: $($JobOutput.Values | ft | Out-String)"
        
        #Create subject depending on processing status
        if ($ErrorCount -ge 1) {
            $Subject = "Employee Transfer - TechID#: $($ADAccount.description) - Processing ERROR"
        }
        elseif ($NewOUPath.Count -gt 1) {
            if ($ErrorCount -ge 1) {
               $Subject = "Employee Transfer - TechID#: $($ADAccount.description) - Processing ERROR"
            }
            else {
                $Subject = "Employee Transfer - TechID#: $($ADAccount.description) - OU Move Required to Complete"
            }
        }
        else {
            $Subject = "Employee Transfer - TechID#: $($ADAccount.description) - Completed"
        }
        
        #Send email
        $to = ""
        $from = ""
        $Subject = $Subject
        $smtpserver = ""
        $body = "
        
        Azure processed employee: $($ADAccount.SamAccountName)
        Job output: $($LocalTranscript | % {$_} | Out-String)
        
        Thank you
        
        
        "
        send-mailmessage -to $to -from $from -Subject $subject -body $body -smtpserver $smtpserver 
        
        Write-Output "Sending report email to $($to)"
    }
}
else {
    Write-Output "There is no Azure variable to run against or too many!"
    exit
}
