# **********************************************************************************
#
# Script Name: Inventory-OpsSystems.ps1
# Version: 1.4
# Author: Aaron O and Dave M
# Date Created: 11-16-2015
# _______________________________________
#
# MODIFICATIONS:
# Date Modified: 11-17-2015
# Modified By: Dave m
# Reason for modification: Modified so it would work on local computer (instead of remotely) so that it would work in DMZ and Workgroup machines.
# What was modified: Replaced several AD calls with WMI calls
#
# Date Modified: 12/4/2015
# Modified By: Dave M
# Reason for modification: Win32_Product calls can cause MSI to "repair" installed apps.
# What was modified: Replaced Win32_Product WMI calls with registry function
#
# Date Modified: 12/21/2015
# Modified By: Dave M
# Reason for modification: Needed more SEP data, RAM, HDD
# What was modified: Added additional data collection.
#
# Date Modified: 2/11/2016
# Modified By: Dave M
# Reason for modification: File cleanup.
# What was modified: Added routines for cleaning up LOG files and CSV files.
#
# Description: Gathers inventory data from the local computer.
#
# Usage:
# ./Inventory-OpsSystems.ps1
#
# **********************************************************************************

<#
	.SYNOPSIS
		Gathers inventory data from the local computer.

	.DESCRIPTION
		Gathers inventory data regarding OS version and editionm CPU, HDD Size, SQL, SEP and Backup Exec versions from the local computer.

	.EXAMPLE
		./Inventory-OpsSystems.ps1

		DESCRIPTION
		===========
		Gathers inventory data from the local computer.
		
	.NOTES
		Requires permissions to access WMI and the Registry.

#>

# Functions and Filters
Function Get-DUninstallablePrograms
{
	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Specify Computer Name(s)")]
		[String[]]
        [Alias("Identity")]
        [Alias("Name")]
        [Alias("SAMAccountName")]
        [Alias("DNSHostName")]
        [Alias("Computer")]
        [Alias("ComputerName")]
        [Alias("Server")]
        [Alias("System")]
        [Alias("Sys")]
$Computers
)

$results = @()
Foreach ($sys IN $Computers)
	{
	#Define the variable to hold the location of Currently Installed Programs
	$UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
	#Create an instance of the Registry Object and open the HKLM base key
	$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$sys)
	#Drill down into the Uninstall key using the OpenSubKey Method
	$regkey=$reg.OpenSubKey($UninstallKey)
	#Retrieve an array of string that contain all the subkey names
	$subkeys=$regkey.GetSubKeyNames()
	#Open each Subkey and use GetValue Method to return the required values for each
	foreach($key in $subkeys)
		{
		$thisKey=$UninstallKey+"\\"+$key
		$thisSubKey=$reg.OpenSubKey($thisKey)
		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $sys
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
		$obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
		$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
		$results += $obj
		} # foreach($key in $subkeys)
	} # Foreach ($sys IN $Computers)
Return $results
} # Function Get-DUninstallablePrograms

# End of Functions and Filters

# Initialization
$error.clear()
$ErrorActionPreference = "SilentlyContinue"
$Sys=gc env:computername
write-host "Discovered $sys, processing..."
[string]$datecode=get-date -DisplayHint Date -Format yyyyMMdd
write-host "Using Date Code: $datecode"
$sstring="Password"
$report = @()

$test=test-path I:\20151204.DAT
if ($test -ne $true)
	{
	NET USE I: /delete
	}

$UNC="`\`\InventoryServerName`\Inventory`$"
NET USE $UNC /DELETE
NET USE $UNC $sstring /USER:InventoryServerName\SubDomaininvadmin
if ($? -ne $true) {$UNC="`\`\InventoryServerName`.SubDomain`.Domain`\Inventory`$";NET USE $UNC /DELETE;NET USE \\InventoryServerName.SubDomain.Domain\Inventory$ $sstring /USER:InventoryServerName\SubDomaininvadmin}
if ($? -ne $true) {$UNC="`\`\192`.168`.0`.2`\Inventory`$";NET USE $UNC /DELETE;NET USE \\10.85.30.175\Inventory$ $sstring /USER:InventoryServerName\SubDomaininvadmin}

$test=test-path $UNC\20151204.DAT
if ($test -ne $true) {write-error "Unable to locate Inventory file repository, files will be stored at c:\Scripts." -erroraction silentlycontinue}

# Used to direct the DS Searcher at a Domain (as opposed to a local directory).
$forestName = ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Name
$ADsPath = [ADSI]"LDAP://$forestName"
write-host "Directory Forest Entry initialized:"
$forestName | fl

# Initializes the DS Searcher
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
write-host "Directory Searcher initialized:"
$objSearcher | fl

# Used to filter on Compute objects.
$strCategory = "Computer"
# Used to isolate server objects
$strOS = "*Windows*"
# Sets the filter for the search
$objSearcher.Filter = "(&(name=$sys)(objectCategory=$strCategory))"

write-host "Loading Properties:"
# Loads the properties that AD will return for each object.
$colProplist = @("DNSHostName","distinguishedname","ManagedBy")
foreach ($i in $colPropList)
	{
write-host "Loading $i"
	$objSearcher.PropertiesToLoad.Add($i) | Out-Null
	}

write-host "DS Searcher properties:"
$objSearcher | fl

#Executes AD Query
$colResults = $objSearcher.FindAll()

write-host "Found $($colResults.count) results."
$colResults.Properties | fl

if ($colResults.count -ne 1)
	{
write-host "Manually adding local computer."
	[array]$colResults=@()
	$PSOResult = New-Object PSObject
	$hashtable=@{}
	$hashtable.Add("dnshostname", "$([System.Net.Dns]::GetHostEntry($(gc env:computername)).HostName)")
	$hashtable.Add("ManagedBy", $null)
	$hashtable.Add("distinguishedname", "WINNT://$(gc ENV:Computername)")
	$PSOResult | Add-Member NoteProperty Properties $hashtable
write-host "Manual results:"
$PSOResult.Properties | fl
	$colResults = $PSOResult
	}

write-host "Processing results..."
# Processes each object returned for reporting.
foreach ($objResult IN $colResults)
	{
write-host "Working on $objResult"
$objResult | fl
# Loads the Object into a variable for processing
	$objComputer = $objResult.Properties
$objResult.Properties | fl
# Creats a Custom PSObject so we can capture the report details.
	$temp = New-Object PSObject
# Captures the DNS Host Name of the object
	$name=$objComputer.dnshostname
	write-host "Processing $name"
	$temp | Add-Member NoteProperty DNSHostName $($objcomputer.dnshostname)
$temp | fl
# Captures the Domain for this server.
	$temp | Add-Member NoteProperty Domain $((gwmi Win32_ComputerSystem).Domain)
$temp | fl
# Captures the IP Address(es) for this server.
	$temp | Add-Member NoteProperty IPAddress $($IPAddress=@();get-WmiObject Win32_NetworkAdapterConfiguration | where {$_.IPAddress -ne $null} | foreach {$IPAddress+=$_.IPAddress};$IPAddress | out-string)
$temp | fl
# Captures the SCOM Management Server for this server.
	$temp | Add-Member NoteProperty SCOM $($objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $name);$objRegKey= $objReg.OpenSubKey("SOFTWARE\\Microsoft\\Microsoft Operations Manager\\3.0\\Agent Management Groups\\EAD\\Parent Health Services\\0" );$objRegKey.GetValue("NetworkName"))
$temp | fl
# Captures the system model of the server
	$temp | Add-Member NoteProperty SystemModel $((gwmi Win32_ComputerSystem).model)
$temp | fl
# Captures the Hyper-V Host for this server.
	$temp | Add-Member NoteProperty PhysicalHostname $($objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $(gc env:computername));$objRegKey= $objReg.OpenSubKey("SOFTWARE\\Microsoft\\Virtual Machine\\Guest\\Parameters" );$objRegKey.GetValue("PhysicalHostNameFullyQualified"))
$temp | fl
# Captures the Managed By Data stored in AD
	$temp | Add-Member NoteProperty ManagedBy $($objcomputer.ManagedBy)
write-host "ManagedBy value: $($objcomputer.ManagedBy)"
$temp | fl
# Captures the AD Description for this server.
	$temp | Add-Member NoteProperty Description $((Get-WmiObject -class Win32_OperatingSystem).Description)
write-host "Description value: $((Get-WmiObject -class Win32_OperatingSystem).Description)"
$temp | fl
# Captures the Role from AD DN for this server.
	$temp | Add-Member NoteProperty PossibleRole $($array=$objcomputer.distinguishedname.split(",");$array[1].replace("OU=",""))
$temp | fl
# Captures the Distinguished Name of the object
	$temp | Add-Member NoteProperty DistinguishedName $($objcomputer.distinguishedname)
write-host "DistinguishedName value: $($objcomputer.distinguishedname)"
$temp | fl
# Queries the SEP version installed on the server.
	$temp | Add-Member NoteProperty SEPversion $((get-itemproperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate').DeployRunningVersion)
$temp | fl
# Queries the Registry for SEP update server info.
	$temp | Add-Member NoteProperty SEPServer $((get-itemproperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate').LastServerIP)
$temp | fl
# Queries the Registry for SEP update time info.
	$temp | Add-Member NoteProperty SEPLastUpdate $((get-itemproperty 'HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate').LatestVirusDefsDate)
$temp | fl
# Queries the CPU Model Name installed on the server.
	$temp | Add-Member NoteProperty CPUName $((gwmi Win32_Processor).name | select -first 1)
$temp | fl
# Captures the Computer Serial Number for this server.
	$temp | Add-Member NoteProperty SerialNumber $((gwmi win32_bios).SerialNumber)
$temp | fl
# Captures the Number of CPUs for this server.
	$temp | Add-Member NoteProperty CPUs $($CPUs=0;(gwmi Win32_Processor) | foreach {$CPUs += $_.NumberOfLogicalProcessors};$CPUs)
$temp | fl
# Captures the Number of CPUs for this server.
	$temp | Add-Member NoteProperty RAMInGB $([math]::Round($((gwmi Win32_ComputerSystem).TotalPhysicalMemory /1GB), 0))
$temp | fl
# Captures the Number of HDDs for this server.
	$temp | Add-Member NoteProperty NumberOfHDDs $((gwmi Win32_logicaldisk  | where {$_.DriveType -eq 3} | measure-object).count)
$temp | fl
# Captures the HDD Space allocated to this server.
	$temp | Add-Member NoteProperty HDDSizeInGB $(gwmi Win32_logicaldisk | where {$_.DriveType -eq 3} | foreach {$Size+=$_.Size}; $size/1GB)
$temp | fl
# Captures the free HDD Space on this server.
	$temp | Add-Member NoteProperty HDDTotalFreeSpaceInGB $(gwmi Win32_logicaldisk | where {$_.DriveType -eq 3} | foreach {$Free+=$_.FreeSpace}; $Free/1GB)
$temp | fl
# Captures the lowest free HDD Space on any single drive on this server.
	$temp | Add-Member NoteProperty HDDLowestFreeSpaceInGB $([uint64]$low=9999999999999;gwmi Win32_logicaldisk | where {$_.DriveType -eq 3} | foreach {if ($_.FreeSpace -lt $low) {$low=$_.FreeSpace}}; $low/1GB)
$temp | fl
# Captures the OS Edition for this server.
	$temp | Add-Member NoteProperty OSEdition $((gwmi Win32_OperatingSystem).caption)
$temp | fl
# Queries the Backup Exec version installed on the server.
	$temp | Add-Member NoteProperty BUXversion $((Get-DUninstallablePrograms -Computers $(gc env:computername) | where {$_.DisplayName -like "*Symantec Backup Exec*"} | select -first 1).DisplayVersion)
$temp | fl
# Queries the SQL Version installed on the server.
	$temp | Add-Member NoteProperty SQLVersion $($inst=(get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances;if ($inst.count -gt 1) {$inst=$inst | select -last 1};$Key = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$inst;(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$Key\Setup").Version)
$temp | fl
# Queries the SQL Edition installed on the server.
	$temp | Add-Member NoteProperty SQLEdition $($inst=(get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances;if ($inst.count -gt 1) {$inst=$inst | select -last 1};$Key = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$inst;(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$Key\Setup").Edition)
$temp | fl
# Adds report time to the report.
	$temp | Add-Member NoteProperty ReportDate $(get-date -DisplayHint Date -Format MM-dd-yy | out-string)
$temp | fl
# Adds Uptime in days to the report.
	$temp | Add-Member NoteProperty UptimeInDays $($wmi = gwmi Win32_OperatingSystem ;($wmi.ConvertToDateTime($wmi.LocalDateTime) - $wmi.ConvertToDateTime($wmi.LastBootUpTime)).Days)
$temp | fl
# Adds the results to the report collection
	$report += $temp
	}
# Returns Report data.
$report | fl
$report | export-csv c:\Scripts\$sys-$datecode.CSV -notypeinformation
$error | out-string > c:\Scripts\ERRORS-$sys-$datecode.LOG

# Copy files
copy-item c:\Scripts\$sys-$datecode.CSV $UNC -force -confirm:$false
if ($? -eq $true)
# Clean up files.
	{
	$Oldfiles=gci c:\Scripts\*.CSV | where {$_.LastWriteTime -lt $((get-date).adddays(-4))}
	$Oldfiles | Remove-Item -force -confirm:$false
	}
copy-item c:\Scripts\ERRORS-$sys-$datecode.LOG $UNC\Logs -force -confirm:$false
if ($? -eq $true)
# Clean up files.
	{
	$Oldfiles=gci c:\Scripts\*.LOG | where {$_.LastWriteTime -lt $((get-date).adddays(-2))}
	$Oldfiles | Remove-Item -force -confirm:$false
	}
copy-item $UNC\Inventory-OpsSystems.ps1 c:\Scripts\  -force -confirm:$false
copy-item $UNC\Inventory-OpsSystems-PSv2.ps1 c:\Scripts\  -force -confirm:$false
if ((get-host).version.Major -lt 3)
	{
	Del c:\Scripts\Inventory-OpsSystems.ps1 -force -confirm:$false
	Start-Sleep -s 1
	Rename-Item -path c:\Scripts\Inventory-OpsSystems-PSv2.ps1 -NewName Inventory-OpsSystems.ps1 -force -confirm:$false
	}
copy-item $UNC\Test.DAT c:\Scripts\  -force -confirm:$false
copy-item $UNC\ManualPatch.CSV c:\Scripts\  -force -confirm:$false

# WSUS Registry monitring and remediation
if ($forestName -eq $null) {$forestName = "WorkGroup"}
if ($forestName.contains("test") -eq $true) {$prod=$false} else {$prod=$true}

if ((Test-Path c:\Scripts\ManualPatch.CSV) -eq $true)
{
[string]$MPL=gc c:\Scripts\ManualPatch.CSV
[bool]$manual=$MPL.contains("$sys")

$WUKey="SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate"
$WUPath="HKLM:\\" + $WUKey
$AUKey="SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate"
$AUPath="HKLM:\\" + $AUKey
$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$sys)
$regkey=$reg.OpenSubKey($WUKey,$true)
if ($prod -ne $true)
{
if ($regkey.GetValue("WUServer") -ne "http://LabWSUSServer.SubSubDomain.SubDomain.test:8530") {$regkey.SetValue("WUServer","http://LabWSUSServer.SubSubDomain.SubDomain.test:8530")}
if ($regkey.GetValue("WUStatusServer") -ne "http://LabWSUSServer.SubSubDomain.SubDomain.test:8530") {$regkey.SetValue("WUStatusServer","http://LabWSUSServer.SubSubDomain.SubDomain.test:8530")}
} else
{
if ($regkey.GetValue("WUServer") -ne "http://WSUSServer.SubDomain.Domain:8530") {$regkey.SetValue("WUServer","http://WSUSServer.SubDomain.Domain:8530")}
if ($regkey.GetValue("WUStatusServer") -ne "http://WSUSServer.SubDomain.Domain:8530") {$regkey.SetValue("WUStatusServer","http://WSUSServer.SubDomain.Domain:8530")}
} # if ($prod -ne $true)

if (($regkey.GetValue("AcceptTrustedPublisherCerts")) -eq $null)
	{
	New-ItemProperty -Path $WUPath -Name "AcceptTrustedPublisherCerts" -Value 1 -PropertyType DWORD -Force
	} # if (($regkey.GetValue("AcceptTrustedPublisherCerts")) -eq $null)

if ($regkey.GetValue("AcceptTrustedPublisherCerts") -ne 1) {$regkey.SetValue("AcceptTrustedPublisherCerts",1)}


$regkey=$reg.OpenSubKey($AUKey,$true)
if ($manual -eq $true)
{
if ($regkey.GetValue("UseWUServer") -ne 0) {$regkey.SetValue("UseWUServer",0)}
If ($regkey.GetValue("AUOptions") -ne 3) {$regkey.SetValue("AUOptions",3)}
} else
{
if ($regkey.GetValue("UseWUServer") -ne 1) {$regkey.SetValue("UseWUServer",1)}
If ($regkey.GetValue("AUOptions") -ne 4) {$regkey.SetValue("AUOptions",4)}
} # If ($regkey.GetValue("AUOptions") -ne 3) {$regkey.SetValue("AUOptions",3)}

if ($regkey.GetValue("AUPowerManagement") -ne 1) {$regkey.GetValue("AUPowerManagement",1)}
if ($regkey.GetValue("IncludeRecommendedUpdates") -ne 1) {$regkey.GetValue("IncludeRecommendedUpdates",1)}
if ($regkey.GetValue("NoAutoUpdate") -ne 0) {$regkey.GetValue("NoAutoUpdate",0)}
if ($regkey.GetValue("ScheduledInstallDay") -ne 0) {$regkey.GetValue("ScheduledInstallDay",0)}
if ($regkey.GetValue("ScheduledInstallTime") -ne 5) {$regkey.GetValue("ScheduledInstallTime",5)}
if ($regkey.GetValue("DetectionFrequencyEnabled") -ne 1) {$regkey.GetValue("DetectionFrequencyEnabled",1)}
if ($regkey.GetValue("DetectionFrequency") -ne 2) {$regkey.GetValue("DetectionFrequency",2)}
if ($regkey.GetValue("NoAutoRebootWithLoggedOnUsers") -ne 0) {$regkey.GetValue("NoAutoRebootWithLoggedOnUsers",0)}
} # if ((Test-Path c:\Scripts\ManualPatch.CSV) -eq $true)

NET USE $UNC /DELETE
