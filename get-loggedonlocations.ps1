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
$EmailFormat = $EmailFormat + "TH{border-width: 1px;padding: 5px 10px 5px 10px;border-style: solid;border-color: black;background-color:royalblue}"
$EmailFormat = $EmailFormat + "TD{border-width: 1px;padding: 5px 10px 5px 10px;border-style: solid;border-color: black;background-color:gainsboro}"
$EmailFormat = $EmailFormat + "</style>"
#
[Hashtable]$LogonType = @{ 0 = "System Account"; 2 = "Interactive"; 3 = "Network"; 4 = "Batch"; 5 = "Service"; 6 = "Proxy"; 7 = "Unlock"; 8 = "NetworkCleartext"; 9 = "NewCredentials"; 10 = "RemoteInteractive"; 11 = "CachedInteractive"; 12 = "CachedRemoteInteractive"; 13 = "CashedUnlock"}
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_HH-mm-ss
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
# This is used for testing to limit the number of servers that are tested.  Adjust the value for $tempcount - lt XXXX
ForEach($tempserver in $allservers){
    $tempcount ++
    IF($tempcount -lt 10000){$servers += $tempserver}
}
#
$allserverscount = $allservers.count
$ServerCount = $servers.count
Write-Host "Found $ServerCount Servers of $allserverscount"
$results = @()
$OfflineServers = @()
$AllUserProcessesResults = @()
$Progress = 0
Start-Sleep -Seconds 5
clear-host
# $servers = Get-ADComputer "WFM-FIM-SYNC-01"
$TempLogonSessions = @()
ForEach ($srv in $servers){
    $srvname = $srv.Name
    Write-Progress -Activity "Searching Servers for Logged On Users" -Status "Progress -- Checking $srvname" -PercentComplete (($Progress/$ServerCount)*100)
    if (test-connection -computername $srv.name -Count 1 -Quiet -ErrorAction SilentlyContinue){
        $TempLogonSessions = Get-WmiObject Win32_LogonSession -ComputerName $srv.name | Where-Object {($_.LogonType -eq '10') -or ($_.LogonType -eq '2')}
        ForEach ($TempLogonSession in $TempLogonSessions) {
            $AllUserProcesses = @()
            $TempSessionID = $TempLogonSession.LogonID
            $TempLoggedOnUser = Get-WmiObject Win32_LoggedOnUser -ComputerName $srvname | Where-Object {($_.Dependent -like "*$TempSessionID*")}
            $TempLoggedOnUserName = $TempLoggedOnUser.Antecedent.split('"')[3]
            $TempLoggedOnUserName = $TempLoggedOnUserName.toUpper()
            $TempLogonType = $LogonType.GetEnumerator() | Where-Object {$_.Name -eq $TempLogonSession.LogonType}
            $TempProcesses = Get-WmiObject Win32_Process -ComputerName $srvName
            $Processes = @()
            ForEach ($TempProcess in $TempProcesses){
                IF(($TempProcess.getowner()).user -eq $TempLoggedOnUserName){
                    $Processes = New-Object psobject
                    $Processes | Add-Member -MemberType NoteProperty -Name "Server" -Value $srvname
                    $Processes | Add-Member -MemberType NoteProperty -Name "Process" -Value $TempProcess.name
                    $Processes | Add-Member -MemberType NoteProperty -Name "Owner" -Value $TempLoggedOnUserName
                    $AllUserProcesses += $Processes
                    $AllUserProcessesResults += $Processes   
                }                
            }
            # IF(!(Get-WmiObject Win32_Process -ComputerName $srvName).getowner() | Select user -Unique | Where-Object {$_.User -eq $TempLoggedOnUserName}) {$TempProcessesRunning = "No Processes Found"} ELSE {$TempProcessesRunning = "Processes Running"}
            IF(!($AllUserProcesses)) {$TempProcessesRunning = "No Processes Running"} Else {$TempProcessesRunning = "Processes Running"}
            IF(!($AllUserProcesses)) {$AllUserProcessesString = "No Processes were found"} Else {$AllUserProcessesString = [system.String]::Join(" | ",$AllUserProcesses.Process)}
            $result = New-Object psobject
            $result | Add-Member -MemberType NoteProperty -Name "Server" -Value $srv.name
            $result | Add-Member -MemberType NoteProperty -Name "Username" -Value $TempLoggedOnUserName
            $result | Add-Member -MemberType NoteProperty -Name "LogonSessionID" -Value $TempSessionID
            $result | Add-Member -MemberType NoteProperty -Name "AuthenticationPackage" -Value $TempLogonSession.AuthenticationPackage
            $result | Add-Member -MemberType NoteProperty -Name "LogonType" -Value $TempLogonType.Value
            $result | Add-Member -MemberType NoteProperty -Name "LogonDay" -Value $TempLogonSession.StartTime.Substring(0,8).Insert(4,'-').Insert(7,'-')
            $result | Add-Member -MemberType NoteProperty -Name "LogonTime" -Value $TempLogonSession.StartTime.Substring(8,13).Insert(2,':').Insert(5,':').split('.')[0]
            $result | Add-Member -MemberType NoteProperty -Name "ProcessesRunning" -Value $TempProcessesRunning
            $result | Add-Member -MemberType NoteProperty -Name "Processes" -Value $AllUserProcessesString 
            $results += $result
           
       }
<# Option 1
ForEach ($srv in $servers){
    $srvname = $srv.Name
    Write-Progress -Activity "Searching Servers for Logged On Users" -Status "Progress -- Checking $srvname" -PercentComplete (($Progress/$ServerCount)*100)
    if (test-connection -computername $srv.name -Count 1 -Quiet -ErrorAction SilentlyContinue){
        $TempLoggedOnUsers = Get-WmiObject Win32_LoggedOnUser -ComputerName $srvname | Where-Object { (($_.Antecedent -like "*WEGMANS*") -and ($_.Antecedent -notlike '.*$'))} | Select-Object -Unique
        # $TempLoggedOnUsers = Get-WmiObject Win32_LoggedOnUser -ComputerName $srv.name | Select Dependent, Antecedent | Where-Object {$_.Antecedent -like "*WEGMANS*"}
        ForEach ($TempLoggedOnUser in $TempLoggedOnUsers){
            $TempLoggedOnUserName = $TempLoggedOnUser.Antecedent.split('"')[3]
            $TempLoggedOnUserSessionID = $TempLoggedOnUser.Dependent.split('"')[1]
            IF($TempLoggedOnUserName -notcontains "$"){
                $TempLogonType = $null
                $TempLoggedOnUserName = $TempLoggedOnUserName.toUpper()
                $TempLoggedOnSessionInfo = Get-WmiObject Win32_LogonSession -ComputerName $srv.name | Where-Object {($_.logonid -match $TempLoggedOnUserSessionID) -and ($_.LogonType -eq '10')}
                $TempLogonType = $LogonType.GetEnumerator() | Where-Object {$_.Name -eq $TempLoggedOnSessionInfo.LogonType}
                $result = New-Object psobject
                $result | Add-Member -MemberType NoteProperty -Name "Server" -Value $srv.name
                $result | Add-Member -MemberType NoteProperty -Name "Username" -Value $TempLoggedOnUserName
                $result | Add-Member -MemberType NoteProperty -Name "LogonSessionID" -Value $TempLoggedOnUserSessionID
                $result | Add-Member -MemberType NoteProperty -Name "AuthenticationPackage" -Value $TempLoggedOnSessionInfo.AuthenticationPackage
                $result | Add-Member -MemberType NoteProperty -Name "LogonType" -Value $TempLogonType.Value
                $result | Add-Member -MemberType NoteProperty -Name "LogonDay" -Value $TempLoggedOnSessionInfo.StartTime.Substring(0,8).Insert(4,'-').Insert(7,'-')
                $result | Add-Member -MemberType NoteProperty -Name "LogonTime" -Value $TempLoggedOnSessionInfo.StartTime.Substring(8,13).Insert(2,':').Insert(5,':').split('.')[0]
                $results += $result            
            }
        }
# Option 1>
        # $TempLoggedOnUsers = Get-WmiObject Win32_LoggedOnUser -ComputerName $srv.name | Select Dependent, Antecedent | Where-Object {$_.Antecedent -like "*WEGMANS*"} | ForEach-Object{$_.Antecedent.split('"')[3]}
        <#
        ForEach ($TempLoggedOnUser in $TempLoggedOnUsers){
            $TempLoggedOnUser = $TempLoggedOnUser.toUpper()
            $result = New-Object psobject
            $result | Add-Member -MemberType NoteProperty -Name "Server" -Value $srv.name
            $result | Add-Member -MemberType NoteProperty -Name "Username" -Value $TempLoggedOnUser
            # $result | Add-Member -MemberType NoteProperty -Name "LogonTime" -Value "Not Currently Determined"
            $results += $result
        }
        $usrProcs = Get-WmiObject Win32_Process -ComputerName $srv.Name -ErrorAction SilentlyContinue | Where {$_.Name -like "explorer.*"}
        ForEach ($proc in $usrProcs){
            $ProcessUser = ($proc.GetOwner().User)
            $ProcessUser = $ProcessUser.toUpper()
            $result = New-Object psobject
            $result | Add-Member -MemberType NoteProperty -Name "Server" -Value $proc.PSComputerName
            $result | Add-Member -MemberType NoteProperty -Name "Username" -Value $ProcessUser
            $result | Add-Member -MemberType NoteProperty -Name "LogonTime" -Value $proc.CreationDate
            $results += $result
        }#>
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
    Else {
        $OfflineServersTemp = New-Object PSObject
        $OfflineServersTemp | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $srvname
        $OfflineServers += $OfflineServersTemp
    }
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
    # $TechAdminResult | ConvertTo-Html Server, Username, LogonSessionID, AuthenticationPackage, LogonType, LogonDay, LogonTime -Head $EmailFormat -Title "Server Logon Report for $RptUsername " -body "$HTMLHead<H2> The Account $RptUsername is logged on to the following servers</H2>The email address for the manager of this account is $TempTo" | Set-Content $OutputFile
    # $To = $ManagerEmail.EmailAddress
    $ManagerEmailAddress = $ManagerEmail.EmailAddress
    $to = "john.shelton@wegmans.com"
    $Subject = "Server Logon Report for $RptUsername"
    $Body = $TechAdminResult | ConvertTo-Html Server, Username, LogonSessionID, AuthenticationPackage, LogonType, LogonDay, LogonTime, ProcessesRunning, Processes -Head $EmailFormat -Title "Server Logon Report for $RptUsername" -body "$HTMLHead<H2> The Account $RptUsername is logged on to the following servers</H2>The email address for the manager of this account is $ManagerEmailAddress" | Out-String
    Send-MailMessage -From $from -To $To -bcc "john.shelton@wegmans.com" -SmtpServer $SmtpServer -Subject $Subject -BodyAsHtml $Body
    $Body | Set-Content $OutputFile
}
$OutputExcel = $path + $FileName + "ServersOffline_" + $ExecutionStamp + ".xlsx" 
$results | Export-Excel -Path $OutputExcel -WorkSheetname "AllLoggedOnSessions" -TableName "AllLoggedOnSessions" -AutoSize
$AllUserProcessesResults | Export-Excel -Path $OutputExcel -WorkSheetname "AllProcesses" -TableName "AllProcessesTBL" -AutoSize 
IF(!$OfflineServers) {Write-Host "No Offline Servers"} Else {$OfflineServers | Export-Excel -Path $OutputExcel -WorkSheetname "OfflineServers" -TableName "TBLOfflineServers" -AutoSize}