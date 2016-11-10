# **********************************************************************************
#
# Script Name: Merge-InventoryData.ps1
# Version: 1.0
# Author: Dave M
# Date Created: 12/9/15
# _______________________________________
#
# MODIFICATIONS:
# Date Modified: N/A
# Modified By: N/A
# Reason for modification: N/A
# What was modified: N/A
#
# Description: Collects and merges inventory data and converts it to Excel.
#
# Usage:
# ./Merge-InventoryData.ps1
#
# **********************************************************************************

<#
	.SYNOPSIS
		Collects and merges inventory data and converts it to Excel.

	.DESCRIPTION
		Collects inventory data from InventoryServerName.SubDomain.Domain and merges inventory data into a combined CSV. The n it converts it to Excel.

	.EXAMPLE
		./Merge-InventoryData.ps1

		DESCRIPTION
		===========
		Collects and merges inventory data and converts it to Excel.
		
	.NOTES
		

#>


# Initlization
# Stopping transcript in case it was not stopped properly last time this script was ran.
write-host "Stopping transcription in case it was not properly stopped before. An error here is normal."
Import-Module Daves -erroraction stop -warningaction silentlycontinue
Import-Module Combine-CSV -erroraction stop -warningaction silentlycontinue
$string="Password"
$SLog=Get-DPassword C:\Inventory\PS.LOG -nocrypt
[System.Security.SecureString]$sstring=ConvertTo-SecureString -String $string -AsPlainText -Force
$creds=New-Object -TypeName PSCredential -ArgumentList 'InventoryServerName\SubDomaininvadmin',$sstring

Stop-transcript -erroraction silentlycontinue
write-host "Transcription logging stopped. Errors after this point are not normal."
$error.clear()
[string]$datecode=get-date -DisplayHint Date -Format yyyyMMdd
# Start Transcript will keep a log of all items output to screen.
start-transcript "C:\Inventory\RAW Data\Log-Merge-$datecode.LOG"
[string]$datecode=get-date -DisplayHint Date -Format yyyyMMdd

New-PSDrive -Name S -Root \\InventoryServerName.SubDomain.Domain\Inventory$ -PSProvider FileSystem -Credential $creds

# Retrieve data from InventoryServerName
write-output "Copying files"
Copy-Item "S:\*-$datecode.csv" "C:\Inventory\RAW Data" -force -confirm:$false
if ($? -eq $true)
	{
	Move-Item "S:\*-$datecode.csv" "S:\Archives" -force -confirm:$false -erroraction silentlycontinue
	write-output "Moving Copied Files to Archive"
	}
# If the copy succeeds, backup the inventory data.

# Merge all the CSVs into a single file.
write-output "Merging files"
Merge-CSV -SourceFolder "C:\Inventory\RAW Data" -Filter *-$datecode.csv -ExportFileName Merged-$datecode.CSV
# Move that file where we can work with it.
write-output "Copying merged file"
Copy-Item "C:\Inventory\RAW Data\Merged-$datecode.CSV" "C:\Inventory\" -force -confirm:$false
Start-Sleep -s 1

write-output "Starting Excel"
$excel = New-Object -comobject Excel.Application
# Set this to true for debugging, false for normal operations.
# $excel.Visible = $True
$excel.Visible = $False
# WKS is a COM object we can use to edit the CSV file in Excel
write-output "Opening file"
$wks = $excel.Workbooks.Open("C:\Inventory\Merged-$datecode.csv")
# This opens the first sheet/tab in Excel and assigns it to another COM object
write-output "Selecint Worksheet"
$Sheet = $Excel.Worksheets.Item(1)

# This selects the entire worksbook and autofist the column width
write-output "Selecting used range"
$Wkb = $Sheet.UsedRange
# Centers header row
$sheet.Cells.Item(1,1).EntireRow.HorizontalAlignment = -4108
# Bolds header row
$sheet.Cells.Item(1,1).EntireRow.Font.Bold=$True
# Adds a 20th header.
$sheet.Cells.Item(1,26).EntireColumn.NumberFormat = "m/d/yy"
# Centers all cells vertically
$sheet.cells.VerticalAlignment = -4108
# Centers the C column
$sheet.Cells.Item(1,3).EntireColumn.HorizontalAlignment = -4108
# Centers the P column
$sheet.Cells.Item(1,16).EntireColumn.HorizontalAlignment = -4108
# Centers the Q column
$sheet.Cells.Item(1,17).EntireColumn.HorizontalAlignment = -4108
# Centers the R column
$sheet.Cells.Item(1,18).EntireColumn.HorizontalAlignment = -4108
# Centers the S column
$sheet.Cells.Item(1,19).EntireColumn.HorizontalAlignment = -4108
# Centers the T column
$sheet.Cells.Item(1,20).EntireColumn.HorizontalAlignment = -4108
# Centers the U column
$sheet.Cells.Item(1,21).EntireColumn.HorizontalAlignment = -4108
# Centers the Z column
$sheet.Cells.Item(1,26).EntireColumn.HorizontalAlignment = -4108
# Centers the AA column
$sheet.Cells.Item(1,27).EntireColumn.HorizontalAlignment = -4108
# Adds commas and reduces decimal place to one
$sheet.Cells.Item(1,19).EntireColumn.NumberFormat = "#,#.0"
# Adds commas and reduces decimal place to one
$sheet.Cells.Item(1,20).EntireColumn.NumberFormat = "#,#.0"
# Adds commas and reduces decimal place to one
$sheet.Cells.Item(1,21).EntireColumn.NumberFormat = "#,#.0"

# Splits the first column
$sheet.application.activewindow.splitcolumn = 1
# Splits the first row
$sheet.application.activewindow.splitrow = 1
# Freezes panes
$sheet.application.activewindow.freezepanes = $true

# Adjusts used range for the new date column
$Wkb = $Sheet.UsedRange
# Sets width to max
$Wkb.EntireColumn.ColumnWidth = 255
# Sets width to autofit
[void]$Wkb.EntireColumn.AutoFit()
# Sets height to autofit
[void]$Wkb.EntireRow.AutoFit()

# This saves the file as a XLSX file
write-output "Saving file"
$wks.SaveAs("C:\Inventory\Merged-$datecode.xlsx",51)

# Close this instance of Excel
write-output "Quitting Excel"
$excel.quit()
# Variable cleanup to free memory
write-output "Cleaning up memory"
Remove-Variable wkb -erroraction silentlycontinue
Remove-Variable sheet -erroraction silentlycontinue
Remove-Variable wks -erroraction silentlycontinue
Remove-Variable excel -erroraction silentlycontinue
Stop-Process -Name Excel -erroraction silentlycontinue
write-output "Execution complete"
Get-process | where {$_.name -like "*excel*"} | stop-process -Force
# Remove drive mapping
write-output "Removing drive mapping"
Remove-PSDrive -Name S -Force -Confirm:$false

# Make sure the Web Client Service is running. This service is required to map a UNC path to a SharePoint URL
$WebClient=Get-Service WebClient
if ($WebClient.Status -ne "Running")
	{
	Start-Service WebClient
	$successful=$?

if (-not $successful)
	{
	$LOG="Error starting WebClient Service, exting..."
write-output("$LOG")
	write-error ("Error starting WebClient Service, exting...") -erroraction stop
	} # if (-not $successful)
	} # if ($WebClient.Status -ne "Running")

# Setup the network service that lets you map drives to SP URLs
$drive = $(New-Object -Com WScript.Network)

# If I: drive is mapped, remove it, we don't know where it points
$s=test-path ("I:\")
if ($s -eq $true)
{
$drive.RemoveNetworkDrive("I:",$true,$true)
$LOG="`$drive.RemoveNetworkDrive(`"I:`",`$true,`$true) Successful: " + $successful
write-output("$LOG")
}

$s=test-path ("I:\")
if ($s -eq $false)
{
# This maps the I: drive to https://SharePointServerName.SubDomain.Domain/windows/servers/Shared Documents/OCIO Migration/Inventory using the svc.powershell.SubDomain account
$drive.MapNetworkDrive("I:","\\SharePointServerName.SubDomain.Domain\windows\servers\Shared Documents\Inventory" , $false , "SubDomain.Domain\svc.powershell.SubDomain", $SLog)
$LOG="`$drive.MapNetworkDrive(`"I:`",`"\\SharePointServerName.SubDomain.Domain\windows\servers\Shared Documents\Inventory`" , `$false , `"SubDomain.Domain\svc.powershell.SubDomain`", `$SLog)" + $successful
write-output("$LOG")
} else
{
# Remove and remap the drive if it still exists
$drive.RemoveNetworkDrive("I:",$true,$true)
$LOG="`$drive.RemoveNetworkDrive(`"I:`",`$true,`$true) Successful: " + $successful
write-output("$LOG")
$s=test-path ("I:\")
if ($s -eq $true)
{
write-error "Error disconnecting I: Drive, exiting..." -erroraction stop
}
$drive.MapNetworkDrive("I:","\\SharePointServerName.SubDomain.Domain\windows\servers\Shared Documents\OCIO Migration\Inventory" , $false , "SubDomain.Domain\svc.powershell.SubDomain", $SLog)
$LOG="`$drive.MapNetworkDrive(`"I:`",`"\\SharePointServerName.SubDomain.Domain\windows\servers\Shared Documents\OCIO Migration\Inventory`" , `$false , `"SubDomain.Domain\svc.powershell.SubDomain`", `$SLog)" + $successful
write-output("$LOG")
}

# Test to see if drive mapped successfully
$docs=test-path ("I:\")
if ($docs)
	{
# Cleanup the old file
	Remove-Item "I:\Inventory.xlsx" -Force -confirm:$false -erroraction silentlycontinue
	$LOG="Remove-Item `"I:\Inventory.xlsx`" -Force -confirm:`$false -erroraction silentlycontinue Successful: " + $successful
write-output("$LOG")
	# $logger.send($LOG)
# Upload the new file
	copy-item -path "C:\Inventory\Merged-$datecode.xlsx" -destination "I:\Inventory.xlsx" -force -confirm:$false
	$LOG="copy-item -path `"C:\Inventory\Merged-$datecode.xlsx`" -destination `"I:\Inventory.xlsx`" -force -confirm:`$false Successful: " + $successful
write-output("$LOG")
	} else
	{
	# $logger.send($error)
	$LOG="Error updating SharePoint, exiting..."
write-output("$LOG")
	write-error ("Error updating SharePoint, exiting...") -erroraction stop
	} # if ($docs)



# Error log for troubleshooting.
$log=$error | out-string
write-output "`nError Log:"
write-output "$log"

Remove-Variable log -erroraction silentlycontinue

# Stop the transcript
Stop-transcript
