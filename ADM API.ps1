
#Create AD user
####################################################################
$ServerPort = ""
$AuthToken = ""
$Domain = ""
$GivenName = "API"
$Surname = "Test3"
$Password = ""
$Description = "API2"
$TemplateName = "Dental Technician Trainee"
$UserURIString = """givenName"":""$GivenName"", ""password"":""$Password"", ""sn"":""$Surname"", ""description"":""$Description"", ""templateName"":""$TemplateName"""
$CreateADUserRequest = "http://"+$ServerPort+"/RestAPI/CreateUser?domainName="+$Domain+"&AuthToken="+$AuthToken+"&inputFormat=[{"+$UserURIString+"}]"

$result = Invoke-WebRequest -Uri $CreateADUserRequest -Method POST 
$result = $result | ConvertFrom-Json
$result.status

#####################################################################








$array = ("Chris,Reese,1010101,Template ppp,E1",
"Chris,Reese,12345,thing 1,E1",
"Chris,Reese,65412,thing 2,",
"Chris,Reese,656565,thing 3,E1"
)

$FinishedResult = @()

#Filter input into array
$array | % {
    $SplitObj = $_.split(",")
    $obj = [pscustomobject] @{
        First = $SplitObj[0]
        Last = $SplitObj[1]
        Description = $SplitObj[2]
        Template = $SplitObj[3]
        LicenseType = $SplitObj[4]
    }
    $FinishedResult += $obj
}


#AD create only group
$DoNotCreateEmailGroup = $FinishedResult | ? {[string]::IsNullOrEmpty($_.licensetype)}