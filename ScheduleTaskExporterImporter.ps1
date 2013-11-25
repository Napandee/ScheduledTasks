Clear-Host
#Global Variables Section

$Global:ScheduledService = $Null
$Global:ScheduledPassword = $Null

#Functions Section
Function ConnectScheduledTaskService($ConnectToHost)
{
$Global:ScheduledService = New-Object -ComObject("Schedule.Service")
$Global:ScheduledService.connect("$ConnectToHost")
}

Function DisconnectScheduledService
{
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Global:ScheduledService)
}

Function CheckDirectory($Check)
{
If(!(Test-Path -Path $Check))
	{
		New-Item $Check -type Directory -Force
	}
}

Function ConvertSecurePassword($SecurePassword)
{
$ConvertPasswordFromSecureString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$Global:ScheduledPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($ConvertPasswordFromSecureString)
}

Function ExportTasks
{
$ExportTasks = $Global:ScheduledService.GetFolder("\").GetTasks(0)
    if ($SaveDirectory -eq $Null)
    {
        $OutputFileTemp = "$Temp\{0}.xml"
    }
    else
    {
        $OutputFileTemp = "$SaveDirectory\{0}.xml"
    }

$ExportTasks | ForEach-Object {
	$Xml = $_.Xml
	$TaskName = $_.Name
	$FinalFile = $OutputFileTemp -f $TaskName
	$Xml | Out-File $FinalFile
    }
}

Function ImportTasks($ImportTask)
{
$ScheduledTaskFolder = $Global:ScheduledService.GetFolder("\")
Get-Item $ImportTask\*.xml | ForEach-Object{
	$ScheduledTaskName = $_.Name.Replace('.xml', '')
	$TaskXmlContent = Get-Content $_.FullName
 	$Task = $Global:ScheduledService.NewTask($null)
 	$Task.XmlText = $TaskXmlContent
 	$ScheduledTaskFolder.RegisterTaskDefinition($ScheduledTaskName, $Task, 6, $ScheduledUser, $Global:ScheduledPassword, 1, $null)  | out-null
    }
}

#Menu Section
$Export = New-Object System.Management.Automation.Host.ChoiceDescription "&Exports Scheduled Tasks", `
    "Exports all Scheduled tasks found on target host"	
$Import = New-Object System.Management.Automation.Host.ChoiceDescription "&Imports Scheduled Tasks", `
	"Imports all scheduled tasks found on target host"
$ExportAndImport = New-Object System.Management.Automation.Host.ChoiceDescription "&Both exports and imports all Scheduled Tasks", `
    "Exports all Scheduled tasks from a host, then import on diffrent host"	
$Title = "Schedule Task Export/Import tool"
$Message = "Welcome to Scheduled Task Exporter/Importer"
$Options = [System.Management.Automation.Host.ChoiceDescription[]]($Export, $Import,$ExportAndImport)
$Result = $host.ui.PromptForChoice($Title, $Message, $Options, 0)

#Switch for choices
switch ($Result)
    {
        0 {
			"Selection: Export scheduled tasks only"
			$Machine = $($Selection = Read-Host "Server to export from, FQDN required. Localhost is Default"
                        if ($Selection) {$Selection} else {'Localhost'})
            $SaveDirectory = Read-Host "Directory to save exported tasks"
            CheckDirectory $SaveDirectory | Out-Null
            "Connecting to $Machine and exporting Scheduled Task, please hold"
            ConnectScheduledTaskService $Machine
            ExportTasks
            "Exported all task from $Machine"
            DisconnectScheduledService | Out-Null
		  }
        1 {
			"Selection: Import scheduled tasks only"
            $Machine = $($Selection = Read-Host "Server to import on, FQDN required. Localhost is Default"
                        if ($Selection) {$Selection} else {'Localhost'})
            $ExportedTasks = Read-Host "Path to XML Files, Multiple files are supported"
            $ScheduledUser = Read-Host "Username for Tasks to be runned under"
            $SecurePassword = Read-Host "Password for user account" -AsSecureString
            "Connecting to $Machine and importing Scheduled Task, please hold"
            ConnectScheduledTaskService $Machine
            ConvertSecurePassword $SecurePassword
            ImportTasks $ExportedTasks
            "Imported all Tasks from $ExportedTasks"
            DisconnectScheduledService | Out-Null
           }
        2 {
            $ErrorActionPreference = "inquire"
            $Temp = "$env:TEMP\TasksTemp\"
            CheckDirectory $Temp | Out-Null
			"Selection: Export and Import scheduled tasks"
            $MachineExport = $($Selection = Read-Host "Server to export from, FQDN required. Localhost is Default"
                        if ($Selection) {$Selection} else {'Localhost'})
            $MachineImport = $($Selection = Read-Host "Server to import on, FQDN required. Localhost is Default"
                        if ($Selection) {$Selection} else {'Localhost'})
            $ScheduledUser = Read-Host "Username for Tasks to be runned under"
            $SecurePassword = Read-Host "Password for user account" -AsSecureString
            "Connecting to $MachineExport and exporting Scheduled Task, please hold"
            ConnectScheduledTaskService $MachineExport
            ExportTasks
            "Exported all tasks from $MachineExport"
            DisconnectScheduledService | Out-Null
            "Connecting to $MachineImport and importing Scheduled Task, please hold"
            ConnectScheduledTaskService $MachineImport
            ConvertSecurePassword $SecurePassword
            ImportTasks $Temp
            "All tasks have been exported from $MachineExport, and imported on $MachineImport, done"
            Remove-Item $Temp -Force -Recurse
            DisconnectScheduledService | Out-Null
		  }
    }