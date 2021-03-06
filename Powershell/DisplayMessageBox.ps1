<#
.SYNOPSIS
    Sample message box and output
.DESCRIPTION
	Sample message box and output
.NOTES
    File Name  : DisplayMessageBox.ps1
    Author     : Janice McCullough
    Requires   : PowerShell Version 3.0
.LINK
	Try it out
   
.EXAMPLE
#>
cls

$objShell = New-Object -ComObject Wscript.Shell
$objShell.Popup("Pop up message from WSCRIPT",0,"Done")

#message box using windows forms
# add the required .NET assembly:
Add-Type -AssemblyName System.Windows.Forms
 
# show the MsgBox:
$result = [System.Windows.Forms.MessageBox]::Show('Do you want to restart?', 'Warning', 'YesNo', 'Warning')
 
# check the result:
if ($result -eq 'Yes')
{
  Restart-Computer -WhatIf
  Write-Warning 'Restart code here'
}
else
{
  Write-Warning 'Skipping Restart'
}

Write-Host ""
Write-Host "Done"
