option explicit

dim sregkey, sregval, sregdata, oReg, username, password
dim soldname, wstype, oLoc, oSvc, stest, swg

'local registry read
soldname = "."
'wscript.echo "winmgmts:{impersonationLevel=impersonate}!\\" & _
'			 soldname & "\root\default\:StdRegProv"
			 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
		soldname & "\root\default:StdRegProv")

sregkey = "SOFTWARE\JSI Telecom\VoiceBox III\Spare"
sregval = "Type"
sregdata = wstype

oReg.GetStringValue HKLM,"SOFTWARE\JSI Telecom\VoiceBox III\Parameters","Server",stest
wscript.echo stest

'now try remote registry

soldname = "AWS100"
'soldname = "192.168.1.101"
swg = "Den"
wstype = "AWS"
UserName = "Administrator"
Password = "vb"

const HKLM = &H80000002
err.Clear
set oLoc = CreateObject("WbemScripting.SWbemLocator")
set oSvc = oLoc.ConnectServer(soldname,"root\default",Username,Password)
'wscript.echo err.Description
set oReg= oSvc.Get("StdRegProv")

wscript.echo "ok"


sregkey = "SOFTWARE\JSI Telecom\VoiceBox III\Spare"
sregval = "WStype"
sregdata = wstype

oReg.GetStringValue HKLM,"SOFTWARE\JSI Telecom\VoiceBox III\Parameters","Server",stest
wscript.echo stest

wscript.echo "done"
wscript.quit