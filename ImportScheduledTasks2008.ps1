Clear-Host
$ErrorActionPreference = "inquire"
#Welcome Message etc
"Welcome to Scheduled Task Importer"
"This script is used to import Scheduled Tasks in the format xml"

#Variables and main functions section
$ExportedTasks = Read-Host "Path to XML Files, Multiple files are supported"
$Machine = $($Selection = Read-Host "Server to import on, FQDN required! Localhost is Default" 
if ($Selection) {$Selection} else {'Localhost'})
$ScheduledUser = Read-Host "Username for Tasks to be runned under!"
$SecurePassword = Read-Host "Password for user account" -AsSecureString

#Main Code
$ScheduledService = New-Object -ComObject("Schedule.Service")
$ScheduledService.connect("$Machine")
$ScheduledTaskFolder = $ScheduledService.GetFolder("\")

$ConvertPasswordFromSecureString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$ScheduledPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($ConvertPasswordFromSecureString)
 
Get-Item $ExportedTasks\*.xml | ForEach-Object{
	$ScheduledTaskName = $_.Name.Replace('.xml', '')
	$TaskXmlContent = Get-Content $_.FullName
 
	$Task = $ScheduledService.NewTask($null)
 
	$Task.XmlText = $TaskXmlContent
 
	$ScheduledTaskFolder.RegisterTaskDefinition($ScheduledTaskName, $Task, 6, $ScheduledUser, $ScheduledPassword, 1, $null)
 
}

"All Scheduled Tasks have been Imported from the following directory -> $ExportedTasks"
#"The following Credentials where supplied and has been configured, $ScheduledUser $ScheduledPassword
"The Scheduled tasks where imported on: $Machine"