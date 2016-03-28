'==========================================================================
'
' VBScript Source File
'
' NAME: change-org.vbs
'
' AUTHOR: Janice McCullough
' DATE  : 04/04/2003
'
' COMMENT: basic test script template to be used frequently
'			
'12==========================================================================

'option explicit

'on error resume next

'18---------- open different object types - set up environment variables

Dim objShell 'Windows Script Host Shell object
Set objShell = CreateObject("WScript.Shell")

Dim objFSO 'Scripting Dictionary object
Set oBJFSO = CreateObject("Scripting.FileSystemObject")

Dim objEnv 'Windows Script Host environment object
Set objEnv = objshell.Environment("Process")


dim sdrive	' drive letter of system drive
dim strcomputer, objWMIService, colItems, objitem, vbdomain, vbsrv

sdrive = objEnv("SYSTEMDRIVE")


'36 --------- get domain name


strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
for each objitem in colitems
	vbdomain=objItem.domain
        vbsrv=objItem.caption
next

'47  ------------ open text file for logging

'sinputfile = "vbusers.inf"
slogfile = "change-org.log"

CONST ForReading=1

'Set objinFile = objFSO.opentextfile(sinputfile, ForReading)
'Set objlogfile =  objfso.createtextfile(slogfile, True)




'60===================== START SCRIPT HERE ====================================

dim smsg, RegOrg, RegOwn, oldown, oldorg,ssearchfor

CONST For_Reading=1
CONST For_Writing=2
CONST For_Appending=8

'68 ------ Read Existing Entries

key="HKEY_LOCAL_MACHINE\software\microsoft\windows nt\currentversion\"

Oldown=objshell.regread(KEY & "registeredowner")
oldorg=objshell.regread(key & "registeredorganization")

'wscript.echo oldown, oldorg


'78 -----

smsg=msgbox("The current Registered Owner is:  " & oldown & vbcrlf & _
	"The Current Registered Organization is:  " & oldorg & vbcrlf & _
	"Do you want to change these?",vbokcancel,"Change Registered Owner/Organization")

if smsg=vbcancel then
	wscript.quit
end if


RegOwn = inputbox("Enter New Registered Owner","RegisteredOwner")
RegOrg = inputbox("Enter New Registered Organization","RegisteredOrganization")


smsg=msgbox("You entered:" & vbcrlf & vbtab & RegOwn & vbcrlf & _
	vbtab & RegOrg & vbcrlf & "CONTINUE?",vbokcancel,"CONTINUE?")


if smsg=vbcancel then
	wscript.quit
end if

ssearchfor="RegisteredOwner"
CALL SEARCH_REGISTRY
'CALL CHANGE_REGISTRY

'ssearchfor="RegisteredOrganization"
'CALL SEARCH_REGISTRY
'CALL CHANGE_REGISTRY

wscript.echo "done"
wscript.quit


'114------------------------ SUBROUTINES -----------------------------------------

SUB SEARCH_REGISTRY

   DIM sregtmp, souttmp, eregline, icnt, sregkey, aregfilelines

'   sregtmp = objshell.environment("process")("Temp") & "\Regtmp.tmp"
    sregtmp = "regtmp.tmp"
   souttmp = "outtmp.tmp"
'    souttmp = objshell.environment("userprofile") & "\desktop & sOuttmp" & _
'		hour(now)& Minute(Now) & ".reg"
'wscript.echo souttmp



if ssearchfor="" then
	wscript.echo "Error - nothing to search for"
	wscript.quit
end if

	' export the registry to a text file
   objshell.run "regedit /e /a " & sregtmp, , True   ' /a enables report as /ansi for XP

'   sjunk=objfso.opentextfile(souttmp,8,True)
   
'      with objfso.getfile(sregtmp)
' 	aregfilelines = split(.openastextstream(1,0).read(.size), vbcrlf)
'      end with

'      for each eregline in aregfilelines
'	if instr(1, eregline, "[",1) >0 then sregkey = eregline
'	If InStr(1, eRegLine, sSearchFor, 1) >  0 Then
'	   If sRegKey <> eRegLine Then
'	      sjunk.WriteLine(vbcrlf & sRegKey) & vbcrlf & eRegLine
'	   Else
'	      sjunk.WriteLine(vbcrlf & sRegKey)
'	   End If
'	   iCnt = iCnt + 1
'	End If
'     Next

   
With objFSO.OpenTextFile(sOutTmp, For_Writing, True)
  .WriteLine("REGEDIT4" & vbcrlf & "; " & WScript.ScriptName & " " & _
    " Janice McCullough" & vbcrlf & vbcrlf & "; Registry search " & _
    "results for string " & Chr(34) & sSearchFor & Chr(34) & " " & Now & _
    vbcrlf & vbcrlf & _
    "; Save the file with a .reg extension to make changes in the registry" & vbcrlf)

  With objFSO.GetFile(sRegTmp)
    aRegFileLines = Split(.OpenAsTextStream(1, 0).Read(.Size), vbcrlf)
  End With

  objFSO.DeleteFile(sRegTmp)

  For Each eRegLine in aRegFileLines
    If InStr(1, eRegLine, "[", 1) > 0 Then sRegKey = eRegLine
    If InStr(1, eRegLine, sSearchFor, 1) >  0 Then
      If sRegKey <> eRegLine Then

        .WriteLine(vbcrlf & sRegKey) & vbcrlf & sNewRegLine
	NewRegKey=mid(sRegKey,2,len(sRegKey)-2) & "\"
	objshell.regwrite NewRegKey & ssearchfor,RegOwn


'wscript.echo "Changing - " & sRegKey & vbcrlf & NewRegKey & vbcrlf & key
''
wscript.echo NewRegKey & ssearchfor & vbcrlf & regown

a=objshell.regread(NewRegKey & ssearchfor)
wscript.echo a 

      Else
        .WriteLine(vbcrlf & sRegKey)
      End If
      iCnt = iCnt + 1
    End If
  Next

  Erase aRegFileLines

  If iCnt < 1 Then
    oWS.Popup "Search completed in " & FormatNumber(Timer - StartTime, 0) & " seconds." & _
              vbcrlf & vbcrlf & "No instances of " & chr(34) & sSearchFor & chr(34) & _
              " found.",, WScript.ScriptName & " " & " Janice McCullough", 4096
    .Close
'    objFSO.DeleteFile(sOutTmp)
    Cleanup()
  End If
  .Close

End With


END SUB



SUB CHANGE_REGISTRY
	wscript.echo "changing registry - " & keypath
END SUB