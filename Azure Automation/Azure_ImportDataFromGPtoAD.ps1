#Grabs all users' data from GP .csv file and updates fields in AD if they exist in the source
#Written by Khang Pham

#Set paths and grab .csv from GP
$strPath=""
$CSV = get-content -literalpath $strPath | convertfrom-csv
$ReportFileName = "$(get-date -f yyyy-MM-dd).csv"
$AllADUsers = Get-ADUser -Filter * -Properties description
$GPtoADSyncResults = @()
$ErrorItem = @()

#Loop through each account and set values if not empty on user account
Foreach ($UserAccount in $CSV) {
    $DetectedChange = 0
    Write-Output "Processing $($useraccount.description)..."

    #Custom object for report
    $obj = [pscustomobject] @{
        Description = $null
        Office = $null
        Title = $null
        Department = $null
        Company = $null
        Manager = $null
    }

    $Description = $UserAccount.description
    $Office = $UserAccount.Office
    $Title = $UserAccount.JobTitle
    $Department = $UserAccount.Department
    $Company = $UserAccount.Company
    $ManagerDescription = $UserAccount.ManagerEmployeeID
    $User = ($AllADUsers | ? {$_.description -eq $Description}).samaccountname

    #If values not empty, modify from GP values
    if ($User -ne $null) {
        try {
            if (![string]::IsNullOrEmpty($Office)) {
                $DetectedChange++
                Set-ADUser $User -Office $Office -ErrorAction Stop #-WhatIf
            }
            if (![string]::IsNullOrEmpty($Title)) {
                $DetectedChange++
                Set-ADUser $User -Title $Title -ErrorAction Stop #-WhatIf
            }
            if (![string]::IsNullOrEmpty($Department)) {
                $DetectedChange++
                Set-ADUser $User -Department $Department -ErrorAction Stop #-WhatIf
            }
            if (![string]::IsNullOrEmpty($Company)) {
                $DetectedChange++
                Set-ADUser $User -Company $Company -ErrorAction Stop #-WhatIf
            }
            if (![string]::IsNullOrEmpty($ManagerDescription)) {
                $ManagerUser = $AllADUsers | ? {$_.description -eq $ManagerDescription}
                if ($ManagerUser -ne $null) {
                    $DetectedChange++
                    Set-ADUser $User -Manager $ManagerUser -ErrorAction Stop #-WhatIf
               }
            }
           
            $UpdatedUser = Get-ADUser $User -Properties Description, Office, Title, Department, Company, Manager
        }
        #If error occurs, email and report
        catch {
            $ErrorItem += $User
            $UpdatedUser = Get-ADUser -Filter {description -eq $Description} -Properties Description, Office, Title, Department, Company, Manager
        }

        $strManagerName = $Null
        
        #Updated manager value
        if (![string]::IsNullOrEmpty($UpdatedUser.Manager)) {
            $strManagerName = ((Get-ADUser -Identity $UpdatedUser.Manager).name)
        }

        #Populate array with updated values
        $obj.Description = $UpdatedUser.Description
        $obj.Office = $UpdatedUser.Office
        $obj.Title = $UpdatedUser.Title
        $obj.Department = $UpdatedUser.Department
        $obj.Company = $UpdatedUser.Company
        $obj.Manager = $strManagerName

        #If change happened, report on it
        if ($DetectedChange -gt 0) {
            #Write-Output "Changed..."
            $GPtoADSyncResults += $obj
        }
    }
}

Write-Output "Errors found: $($ErrorItem.count)"
Write-Output "Error list  : $($ErrorItem | % {$_ | Out-String})"

#Generate logfile
$GPtoADSyncResults | Export-Csv -NoTypeInformation -Path $ReportFileName

#If any errors occurred, send email
if (![string]::IsNullOrEmpty($ErrorItem)) {
    #Error report email
    $to = ""
    $from = ""
    $Subject = "GP to AD Sync Error"
    $smtpserver = ""
    $body = "
Report of failure (Count: $($ErrorItem.count)) on AD to GP sync.
Error list  : $($ErrorItem | % {$_ | Out-String})
    
Path: $($ReportFileName)            
                           
                "
    send-mailmessage -to $to -from $from -Subject $subject -body $body -smtpserver $smtpserver 
}