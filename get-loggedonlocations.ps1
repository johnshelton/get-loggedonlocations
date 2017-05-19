<#
.SYNOPSIS
    Looks for servers where auser is logged on.
.DESCRIPTION
    Was written with an admin account in mind so it limits the search to servers but could be modified to run against
    enabled computers. For each computer that is found a simple online test is initiated if that test finds the computer is
    reachable the current processes are checked. Each instance of explorer is evaluated for the owner. If the username matches the owner
    it is reported. I use this in conjunction with get-lockedoutlocation when trying to find a reason for an account to continue
    to get unlocked. I scoped this to servers as that was more applicable for my use but the get-adcomputer can be changed based
    on need.
.PARAMETER username
    Specifies the logonname to evaluate and report on.
.EXAMPLE
    C:\PS>c:\etc\scripts\get-loggedonlocations.ps1 -username someuser
    Will return results for any instances matching someuser
.NOTES
    Author: John J. Kavanagh
    03.09.2017 JJK: TODO: Create Advanced function out of this script.
    04.06.2017 JJK: TODO: Provide switch param set to pick the computer type to run against
#>
# Param($username)
# $username = "techadminjs2"
#
# Define Output Variables
#
#
# EMail Settings
#
$from = "windows.support@wegmans.com"
$smtpserver = "smtp.wegmans.com"
#
# Email Format Settings
#
$EmailFormat = "<Style>"
$EmailFormat = $EmailFormat + "BODY{background-color:White;}"
$EmailFormat = $EmailFormat + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$EmailFormat = $EmailFormat + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:royalblue}"
$EmailFormat = $EmailFormat + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:gainsboro}"
$EmailFormat = $EmailFormat + "</style>"
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_hh-mm-ss
$path = "c:\temp\"
$FilenamePrepend = 'temp_'
$FullFilename = "get-loggedonlocations.ps1"
$FileName = $FullFilename.Substring(0, $FullFilename.LastIndexOf('.'))
$FileExt = '.html'
#
$PathExists = Test-Path $path
IF($PathExists -eq $False)
    {
    New-Item -Path $path -ItemType  Directory
    }
$servers = @()
$tempcount = 0
Write-Host "Getting a list of all servers that are enabled"
$allservers = get-adcomputer -Filter {OperatingSystem -like "*Windows*Server*" -AND Enabled -eq $True}
ForEach($tempserver in $allservers){
    $tempcount ++
    #$tempcount
    #$tempserver.Name
    IF($tempcount -lt 10000){$servers += $tempserver}
}
$allserverscount = $allservers.count
$ServerCount = $servers.count
Write-Host "Found $ServerCount Servers of $allserverscount"
$results = @()
$OfflineServers = @()
$Progress = 0
Start-Sleep -Seconds 5
clear-host
ForEach ($srv in $servers){
    $srvname = $srv.Name
    Write-Progress -Activity "Searching Servers for Logged On Users" -Status "Progress -- Checking $srvname" -PercentComplete (($Progress/$ServerCount)*100)
    if (test-connection -computername $srv.name -Count 1 -Quiet -ErrorAction SilentlyContinue){
        $usrProcs = Get-WmiObject Win32_Process -ComputerName $srv.Name -ErrorAction SilentlyContinue | Where {$_.Name -like "explorer.*"}
        ForEach ($proc in $usrProcs){
            $ProcessUser = ($proc.GetOwner().User)
            $ProcessUser = $ProcessUser.toUpper()
            $result = New-Object psobject
            $result | Add-Member -MemberType NoteProperty -Name "Server" -Value $proc.PSComputerName
            $result | Add-Member -MemberType NoteProperty -Name "Username" -Value $ProcessUser
            $result | Add-Member -MemberType NoteProperty -Name "LogonTime" -Value $proc.CreationDate
            $results += $result
        }
        <#
        if($usrProcs){            
            ForEach ($proc in $usrProcs){              
                if ($proc.GetOwner().User -eq $username){
                    "User logon found on {0}" -f $srv.Name        
                }           
            }
        #>
    $Progress ++
    }
    Else {$OfflineServers += $srvname}
}
$TechAdminLogOns = $Results | Where-Object {$_.Username -like "TechAdmin*"}
$TechAdminLoggedOnAccounts = $TechAdminLogOns | Select Username | Sort-Object Username | Get-Unique -AsString
ForEach ($TechAdminLoggedOnAccount in $TechAdminLoggedOnAccounts) {
    $RptUsername = $TechAdminLoggedOnAccount.Username
    $OutputFile = $path + $FilenamePrePend + '_' + $FileName + '_' + $RptUsername + '' + $ExecutionStamp + $FileExt
    $TechAdminResult = @()
    ForEach ($TechAdminLogOn in $TechAdminLogOns){
        IF($TechAdminLogOn.Username -eq $TechAdminLoggedOnAccount.Username){$TechAdminResult += $TechAdminLogOn}
    }
    $Manager = (get-aduser (get-aduser $TechAdminLoggedOnAccount.username -Properties manager).manager).samaccountname
    $ManagerEmail = Get-AdUser $Manager -Properties DisplayName, EmailAddress
    $TechAdminResult | ConvertTo-Html Server, Username, LogonTime -Head $EmailFormat -Title "Server Logon Report for $RptUsername whose ma" -body "$HTMLHead<H2> The Account $RptUsername is logged on to the following servers</H2>" | Set-Content $OutputFile
    $To = $ManagerEmail.EmailAddress
    # $to = "john.shelton@wegmans.com"
    $Subject = "Server Logon Report for $RptUsername"
    $Body = $TechAdminResult | ConvertTo-Html Server, Username, LogonTime -Title "Server Logon Report for $RptUsername" -body "$HTMLHead<H2> The Account $RptUsername is logged on to the following servers</H2>The email address for the manager of this account is $TempTo" | Out-String
    Send-MailMessage -From $from -To $To -SmtpServer $SmtpServer -Subject $Subject -BodyAsHtml $Body
}
$OutputExcel = $path + $FileName + "ServersOffline_" + $ExecutionStamp + ".xlsx" 
$OfflineServers | Export-Excel -Path $OutputExcel -TableName "OfflineServers" -WorkSheetname "OfflineServers"
$results | Export-Excel -Path $OutputExcel -TableName "AllLoggedOnSessions" -WorkSheetname "AllLoggedOnSessions"
