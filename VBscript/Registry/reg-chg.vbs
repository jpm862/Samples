Set wmiServices = GetObject("winmgmts:root/default") 
Set wmiSink = WScript.CreateObject("WbemScripting.SWbemSink", "SINK_") 


wmiServices.ExecNotificationQueryAsync wmiSink, _ 
"SELECT * FROM RegistryValueChangeEvent WHERE Hive='HKEY_LOCAL_MACHINE' AND " & _ 
"KeyPath='SOFTWARE\\Microsoft\\Ole' AND ValueName='DefaultAccessPermission'"

'"KeyPath='SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion' AND ValueName='CSDVersion'" 


WScript.Echo "Listening for Registry Change Events..." & vbCrLf 

While(1) 
WScript.Sleep 1000 
Wend 

Sub SINK_OnObjectReady(wmiObject, wmiAsyncContext) 
WScript.Echo "Received Registry Change Event" & vbCrLf & _ 
"------------------------------" & vbCrLf & _ 
wmiObject.GetObjectText_() 
End Sub
