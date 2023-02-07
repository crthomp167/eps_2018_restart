##
## Designed for EPS 2018+ and to be ran from a single, remote server. Preferrably on a server with a service 
## account that has admin access to the EPS servers.  Recommend running from something other than one of the EPS servers.
## 
## Make sure nothing is restarting your EpicPrintServiceXX service too quickly or the script will hang waiting
## for EpicPrintServiceXX and/or spooler to become fully stopped before proceeding.
## 
## This is for restarting the service and flushing stale spooler files. Not for rebooting.  
## You can comment out 'net start spooler' and the  
## 'Set-EPSServiceMode start' line then add your own reboot line

## ---------------------------------------------------------------------------------------------------------------------------------

## Taken from WAM EPS Failover Script

## <#.OUTPUTS
##    The two or three digit integer corresponding to the highest running version of EPS
##    (e.g If Epic November 2020 is the highest running version it will return 95)
## #>
## function Get-HighestEPSMajorVersion
## {
##    $runningEPSServices = Get-Service |
##        Where-Object -FilterScript { ($PSItem.Name -like "EpicPrintService*") -and ($PSItem.Status -eq "Running") } |
##        Select-Object -ExpandProperty Name
##
##    $highestVersion = ($runningEPSServices | ForEach-Object { $PSItem.Replace("EpicPrintService", "") -as [int] } | Measure-Object -Maximum).Maximum
##
##    return $highestVersion
## }

## if (-not $PSBoundParameters.ContainsKey("EPSMajorVersion"))
## {
##     $EPSMajorVersion = Get-HighestEPSMajorVersion
## }

## ---------------------------------------------------------------------------------------------------------------------------------

#$ErrorActionPreference= 'silentlycontinue'
#$WarningAction= 'silentlycontinue'

##  Ex: .\restart_eps_param.ps1 -ENV PRD

param (
        [Parameter(Mandatory = $true, ParameterSetName = 'ENV')]
        [String]$ENV
)

## Array - enter all servers inside the array parentheses - maintain quotes just to be on the safe side
    
if ($ENV -eq 'PRD') {
    $servers = @("EPSPRD1","EPSPRD2","EPSPRD3","EPSPRD4")
}   
elseif ($ENV -eq 'NONPRD') {
    $servers = @("EPSNONPRD1","EPSNONPRD2","EPSNONPRD3","EPSNONPRD4")
}
## You can eliminate the need for this 'else' by making the above parameter Mandatory = $true
## If Mandatory = $true and you don't pass in a parameter you'll never get this far.  It will error immediately.
## You can use this as a jumping off point to write in some more catches
else {
$date = Get-Date
$date = $("SCRIPT FAILED DUE TO NO ENV SELECTED " + $date + "")
$body = $("SCRIPT FAILED DUE TO NO ENV SELECTED " + $date + "")
Send-MailMessage -To "DL@contoso.com" -From noreplyrestart@contoso -Subject $date -Body $body -SmtpServer relay.contoso.com
}

## Below 'if-else' not needed
## You can use this as a jumping off point to write in some more catches
if ($servers -eq $null) {
Exit
}
else {

foreach ($server in $servers) {
Invoke-Command -ComputerName $server -ScriptBlock {
## Import module can be used if you want to pull this out of an "Invoke-Command -ComputerName -ScriptBlock"
#Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Epic.Core.Printing.Deployment\Epic.Core.Printing.Deployment.dll

## Old stop config
## Make sure EpicPrintService8x reflects your current version
## Not sure if the Epic.Core.Printing.Deployment.dll module will accept a wildcard entry
#Set-EPSServiceMode -Servicename EpicPrintService86 -Action Stop
#(Get-Service -Name EpicPrintService86).WaitForStatus("Stopped")
#sleep -Milliseconds 5000

## New stop config
$service = Get-Service EpicPrintService86
if ( $service.Status -eq [ServiceProcess.ServiceControllerStatus]::Running ) {
    Set-EPSServiceMode -Servicename EpicPrintService86 -Action Stop
    sleep -Milliseconds 60000
}

$service = Get-Service EpicPrintService86
if ( $service.Status -eq [ServiceProcess.ServiceControllerStatus]::Running ) {
    $service.Stop()
    sleep -Milliseconds 5000
}

## The next three lines will stop the spooler service, and remove all stale jobs stuck in printer queues. 
net stop spooler
sleep -Milliseconds 5000

## Old Remove-Item line
## Dangerous if ErrorActionPreference is set to default (AKA continue)
# Remove-Item ((Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers).DefaultSpoolDirectory+"\*") -Recurse -Force 

## Below is a semi-catch for a if this ItemProperty \ Directory does not exist.
## This ItemProperty \ Directory will not exist if there are no printers installed. AKA, a fresh, new server.
## Follow guide here to build a nicer reg entry catch - https://devblogs.microsoft.com/scripting/catch-powershell-errors-related-to-reading-the-registry
## to use the above Remove-Item again.  If you don't catch this, it will throw an error, keep moving, and then start deleting everything on C:\   NOT GOOD
$spoolerdir = (Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers).DefaultSpoolDirectory
if ($spoolerdir -eq $null) {
	$date = Get-Date
	$date = $($ENV + " DANGER WILL ROBINSON " + $date + "")
	Send-MailMessage -To "email.address@contoso.com" -From noreplyrestart@domain.org -Subject $date -Body "SpoolerDir equals null" -SmtpServer mailhost.contoso.com
	## Exit because there's no point continuing if this directory does not exist.
	Exit
}
else {

Remove-Item -Path ($spoolerdir+"\*") -Recurse -Force
## Email lines for testing to see if script gets this far.
#$date = Get-Date
#$date = $($ENV + " LINE 155 WORKED " + $date + "")
#Send-MailMessage -To "christopher.thompson2@nm.org" -From noreplyrestart@nm.org -Subject $date -Body "If at 155 worked" -SmtpServer mailhost.nmh.org ## old relay - shp0a04wv.ch.cadhlt.org
# ------------------

sleep -Milliseconds 5000
net start spooler
sleep -Milliseconds 1000

## Start EpicPrintService86 with 
Set-EPSServiceMode -Servicename EpicPrintService86 -Action Start

## Added 10 seconds before moving onto the next server to give this the previous server time to get back into a fully Running state
sleep -Milliseconds 10000
}
}
$date = Get-Date
$date = $($ENV + " EPS RESTART / SPOOLER FLUSH COMPLETED " + $date + "")
$body = $($ENV + " EPS restart and spooler flush of " + $servers + " completed.")
Send-MailMessage -To "DL@contoso.com" -From noreplyrestart@contoso -Subject $date -Body $body -SmtpServer relay.contoso.com
}
