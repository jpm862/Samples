'==========================================================================
'
' VBScript Source File
'
' NAME: TEST.vbs
'
' AUTHOR: Janice McCullough
' DATE  : 04/04/2003
'
' COMMENT: fix DCOM security issue by deleting regkey
'			
'12==========================================================================

'option explicit

'on error resume next

'open registry

const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Ole"

oReg.DeleteValue HKEY_LOCAL_MACHINE, strKeyPath, "DefaultAccessPermission"

wscript.echo "Finished Deleting Default Access Permission"