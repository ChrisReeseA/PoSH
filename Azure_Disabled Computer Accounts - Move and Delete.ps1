
#Move disabled accounts to OU
#Delete disabled/out-of-date from SCCM
#
$TranscriptPath = ""

$Date = Get-Date
$DateFormatted = Get-Date -Format MMM.dd.yyyy-hhmm
$Days30 = '-30'
$Days60 = '-60'
[int]$ErrorCount = 0
$OutofTimeDate30 = $date.AddDays($Days30)
$OutofTimeDate60 = $date.AddDays($Days60)

$OldComputers = Get-ADComputer -filter {(enabled -eq $true) -and (lastlogondate -lt $OutofTimeDate30) -and (operatingsystem -notlike "*server*")} -Properties lastlogondate | ? {$_.name -notlike "*KIOSK*"}
$DisabledComputers = Get-ADComputer -filter {(enabled -eq $false) -and (lastlogondate -lt $OutofTimeDate30) -and (operatingsystem -notlike "*server*")} -Properties lastlogondate 
$OutOfDateComputers = Get-ADComputer -filter {lastlogondate -lt $OutofTimeDate60 -and operatingsystem -notlike "*server*"} -Properties lastlogondate | ? {$_.name -notlike "*KIOSK*"}
$DisabledComputersWithDescription = Get-ADComputer -Filter 'description -like "Disabled*"' -Properties description

$InactiveOUPath = ""
$LogPath = "$($DateFormatted).log"
$LogFile = New-Item -Path $LogPath

#Move to Inactive OU -> 30 days
"++++++++++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
"Old Machines - 30 day move to ou" | Out-File -FilePath $LogPath -Append
"++++++++++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
if (@($OldComputers).count -le 300) {
    $OldComputers | % {
        #Write-Host "Processing [ $($_.name) ]" -ForegroundColor Gray
        if ($_.distinguishedname -like "*Computers_InActive*") {}
        else {
            try {
                "Moving $($_.name) to Inactive OU Path" | Out-File -FilePath $LogPath -Append
                $_ | Move-ADObject -TargetPath $InactiveOUPath -ErrorAction Stop #-whatif 
            }
            catch {
                $ErrorCount++
                "$($_.name) failed!" | Out-File -FilePath $LogPath -Append
            }
        }
    }
}
else {
    $ErrorCount++
    if (![string]::IsNullOrEmpty($OldComputers)) {
        "ERROR: Too many objects to delete... COUNT: $($OldComputers.Count)!" | Out-File -FilePath $LogPath -Append
    }
    else {
        "No computers available to delete" | Out-File -FilePath $LogPath -Append
    }
}

#Move to Inactive OU -> 30 days + disabled
"+++++++++++++++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
"Disabled Machines - 30 day move to ou" | Out-File -FilePath $LogPath -Append
"+++++++++++++++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
if (@($DisabledComputers).count -le 50) {
    $DisabledComputers | % {
        #Write-Host "Processing [ $($_.name) ]" -ForegroundColor Gray
        if ($_.distinguishedname -like "*Computers_InActive*") {}
        else {
            try {
                "Moving $($_.name) to Inactive OU Path" | Out-File -FilePath $LogPath -Append
                $_ | Move-ADObject -TargetPath $InactiveOUPath -ErrorAction Stop #-whatif 
            }
            catch {
                $ErrorCount++
                "$($_.name) failed!" | Out-File -FilePath $LogPath -Append
            }
        }
    }
}
else {
    $ErrorCount++
    if (![string]::IsNullOrEmpty($DisabledComputers)) {
        "ERROR: Too many objects to delete... COUNT: $($DisabledComputers.Count)!" | Out-File -FilePath $LogPath -Append
    }
    else {
        "No computers available to delete" | Out-File -FilePath $LogPath -Append
    }
}

#Remove from AD -> 60 days + non-server
"++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
"Delete from AD - 60 days" | Out-File -FilePath $LogPath -Append
"++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
#Check for too many
if (@($OutOfDateComputers).count -le 50) {
    $OutOfDateComputers | % {
        #Write-Host "Processing [ $($_.name) ]" -ForegroundColor Gray
        try {
            "Removing $($_.name)" | Out-File -FilePath $LogPath -Append
            Remove-ADObject -Recursive -Identity $_.distinguishedname -Confirm:$false -ErrorAction Stop #-whatif
        }
        catch {
            $ErrorCount++
            "$($_.name) failed!" | Out-File -FilePath $LogPath -Append
        }
    }
}
else {
    $ErrorCount++
    if (![string]::IsNullOrEmpty($OutOfDateComputers)) {
        "ERROR: Too many objects to delete... COUNT: $($OutOfDateComputers.Count)!" | Out-File -FilePath $LogPath -Append
    }
    else {
        "No computers available to delete" | Out-File -FilePath $LogPath -Append
    }
}


#Delete disabled machines with disable date description value
"+++++++++++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
"Disabled Machines - 7 day removal" | Out-File -FilePath $LogPath -Append
"+++++++++++++++++++++++++++++++++" | Out-File -FilePath $LogPath -Append
if (@($DisabledComputersWithDescription).count -le 70) {
    $DisabledComputersWithDescription | % {
        Write-Host "Processing [ $($_.name) ]" -ForegroundColor Gray
        [datetime]$ExpireDate = ($_.description).replace("Disabled: ","")
        if ($ExpireDate -le $date.AddDays(-7)) {
            try {
                "Removing $($_.name)" | Out-File -FilePath $LogPath -Append
                Remove-ADObject -Recursive -identity $_.distinguishedname -Confirm:$false #-whatif
            }
            catch {
                $ErrorCount++
                "$($_.name) failed!" | Out-File -FilePath $LogPath -Append
            }
        }
    }
}
else {
    $ErrorCount++
    "ERROR: Too many objects to delete... COUNT: $($DisabledComputersWithDescription.Count)!" | Out-File -FilePath $LogPath -Append
}

$EmailReport = Get-Content $LogPath

#Prepare and send email report
if ($ErrorCount -ge 1) {
    $to = ""
    $from = ""
    $subject = "Workstation Decommission Error Report"
    $smtpserver = ""
    $body = "
    
$($EmailReport | out-string)   
    
    "
    send-mailmessage -to $to -from $from -Subject $subject -body $body -smtpserver $smtpserver 
}
