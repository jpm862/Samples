<#
.SYNOPSIS
    SWITCH statement
.DESCRIPTION
	Sample Switch statement
.NOTES
    File Name  : SwitchStatement.ps1
    Author     : Janice McCullough
    Requires   : PowerShell Version 3.0
.LINK
   
.EXAMPLE
#>

cls
$items = ("blue","red","banana","orange","purple", "plum")

foreach ($item in $items)
{
	switch ($item)
	{
		{"blue","purple" -contains $_} 			{Write-host "This is a colour"}
		"red"    								{Write-host "This is the BEST colour!"}
		{"banana","plum","apple" -contains $_}	{Write-host "This is a fruit"}
		default									{Write-Host "Is it a fruit or a colour????"}
	}
	
}

Write-Host ""
Write-Host "Done"
