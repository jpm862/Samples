'==========================================================================
'
' VBScript Source File
'
' NAME: change-reg.vbs
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

sinputfile = "change-reg.inf"
slogfile = "change-reg.log"

CONST ForReading=1

Set objinFile = objFSO.opentextfile(sinputfile, ForReading)
Set objlogfile =  objfso.createtextfile(slogfile, True)




'60===================== START SCRIPT HERE ====================================

dim smsg, ssearchfor, strstart, strend

CONST For_Reading=1
CONST For_Writing=2
CONST For_Appending=8

smsg=msgbox("This will modify the Registry - Do you want to Continue?", 305, "WARNING")
if smsg=vbcancel then
	wscript.quit
end if


' -------------- read through inf file from beginning to end

dim strnextline, arrNextLine, snewline, sNewRegValue, sregtmp, souttmp

key="HKEY_LOCAL_MACHINE\software\microsoft\windows nt\currentversion\"


' -------------- setup temporary file names

   sregtmp = objshell.environment("process")("Temp") & "\Regtmp.tmp"


'  -------------- open temporary files

    sregtmp = objshell.environment("process")("Temp") & "\Regtmp.tmp"
    souttmp = objshell.environment("Process")("Temp") & "\sOuttmp.tmp"
'    sregtmp = "regtmp.tmp"
'    souttmp = "outtmp.tmp"




	' export the registry to a text file
   objshell.run "regedit /e /a " & sregtmp, , True   ' /a enables report as /ansi for XP


Do Until objinFile.AtEndOfStream

'-- read the next line of the text file

   strNextLine = objinFile.Readline
   arrNextLine = Split(strNextLine , vbcrlf)
   snewline = arrNextLine(0)
   if left(snewline,1)<>";" then

      arrline=split_line(",",snewline)

'-- write the data out to a log file

      objlogfile.Writeline (snewline & vbcrlf & "----------------------------------->")

' ------ Read Regitry Entry


      OldReg=objshell.regread(KEY & arrline(strstart))
      snewline="Key=" & arrline(strstart) & vbtab & "New Value=" & arrline(strend) & _
 	  vbtab & "Old Value=" & OldReg

      ssearchfor=arrline(strstart)
      sNewRegValue=arrline(strend)

      CALL SEARCH_REGISTRY

   else
	' comment line - do nothing
   end if
Loop


    objFSO.DeleteFile(sregtmp)
    objFSO.DeleteFile(sOutTmp)


objinfile.close
objlogfile.close



smsg=msgbox("Registry Changes Completed", vbokonly, "DONE")

WSCRIPT.QUIT



' ---------------------- FUNCTIONS ---------------------------------



' ----- SPLIT LINE FUNCTION ---
'
' split a line of text into an array based upon a given delimeter.
' array is returned as data for function

FUNCTION SPLIT_LINE(patrn,sstring)

   dim arrstring
'objlogfile.writeline( patrn, sstring)
   arrstring=split(sstring,patrn,-1,1)
   strstart=lbound(arrstring)
   strend=ubound(arrstring)
   split_line=arrstring

END FUNCTION




'114------------------------ SUBROUTINES -----------------------------------------

SUB SEARCH_REGISTRY

   DIM eregline, icnt, sregkey, aregfilelines




if ssearchfor="" then
	wscript.echo "Error - nothing to search for"
	wscript.quit
end if

With objFSO.OpenTextFile(sOutTmp, For_Writing, True)
  .WriteLine("REGEDIT4" & vbcrlf & "; " & WScript.ScriptName & " " & _
    " Janice McCullough" & vbcrlf & vbcrlf & "; Registry search " & _
    "results for string " & Chr(34) & sSearchFor & Chr(34) & " " & Now & _
    vbcrlf & vbcrlf & _
    "; Save the file with a .reg extension to make changes in the registry" & vbcrlf)

  With objFSO.GetFile(sRegTmp)
    aRegFileLines = Split(.OpenAsTextStream(1, 0).Read(.Size), vbcrlf)
  End With


  For Each eRegLine in aRegFileLines
    If InStr(1, eRegLine, "[", 1) > 0 Then sRegKey = eRegLine
    If InStr(1, eRegLine, sSearchFor, 1) >  0 Then
      If sRegKey <> eRegLine Then

        .WriteLine(vbcrlf & sRegKey) & vbcrlf & sNewRegLine
	NewRegKey=mid(sRegKey,2,len(sRegKey)-2) & "\"
	objshell.regwrite NewRegKey & ssearchfor,sNewRegValue

	objlogfile.writeline("Changed Registry:  " & NewRegKey & ssearchfor &_
		 vbcrlf & "To:  " & sNewRegValue & vbcrlf)

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
    Cleanup()
  End If
  .Close

End With


END SUB

