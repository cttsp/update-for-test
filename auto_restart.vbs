
' README
' Need to change 2 variables:
' 1. rootDir
' 2. 

Function getDateTime()
Dim dateStr
dateStr = Year(Date) & "-" & Right("00" & Month(Date), 2) & "-" & Right("00" & Day(Date), 2) & "T" & Time & " [INFO] "
getDateTime = dateStr
End Function


Function killProcess()
Dim oShell
Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.Run "taskkill /f /fi ""IMAGENAME eq face_crm_client.exe""", , True
End Function


Function runClient(rootDir)
Set WinScriptHost = CreateObject("WScript.Shell")
WinScriptHost.CurrentDirectory = rootDir
WinScriptHost.Run rootDir + "face_crm_client.exe", 0
Set WinScriptHost = Nothing
End Function

' Run another script to send log 
Function runSendErrorLog(rootDir)
Set WinScriptHost = CreateObject("WScript.Shell")
Dim arg
arg = "today"
WinScriptHost.Run """util\log_storage_client.exe"" ""today""", 0
Set WinScriptHost = Nothing
End Function


' Some contanst variables
rootDir = "C:\ai-crm-client-artifact\"
logDir = rootDir & "logs\"
'checkFile = rootDir & "logs\" & Year(Date) & "_" & Right("00" & Month(Date), 2) & "_" & Right("00" & Day(Date), 2) & ".log"
checkFile = logDir & "ping.log"
'WScript.echo checkFile
logFile = logDir + "cron.log"

' If ping file is not existed
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
If ( Not oFSO.FolderExists(logDir)) Then
	oFSO.CreateFolder  logDir
End If

If (Not oFSO.FileExists(checkFile)) Then
	oFSO.CreateTextFile checkFile, True
End If

If (Not oFSO.FileExists(logFile)) Then
	oFSO.CreateTextFile logFile, True
	runClient(rootDir)
	Wscript.Quit
End If

' Object to write log
' 2 for Writing
dim logMessage
SET fileSys = CREATEOBJECT("Scripting.FileSystemObject")
Set myLog = fileSys.OpenTextFile(logFile, 8, True)
myLog.WriteLine logMessage + getDateTime() + "================================================================="
myLog.WriteLine logMessage + getDateTime() + "START CRON"
myLog.WriteLine logMessage + getDateTime() + "Check file: " + checkFile


' Get last modified time of log file
' Then save to a variable
SET checkFileObj = fileSys.GetFile(checkFile)
lastModified = checkFileObj.DateLastModified
myLog.WriteLine logMessage + getDateTime() + "Last modified: " & lastModified

'sleep
'sleepTime = 60000
sleepTime = 70000
myLog.WriteLine logMessage + getDateTime() + "Sleep in " & sleepTime  & " milis"
WScript.Sleep sleepTime

' recheck last modified
lastModified2 = checkFileObj.DateLastModified
myLog.WriteLine logMessage + getDateTime() + "Last modified: " & lastModified2


' check and runapp
If DateDiff("s", lastModified, lastModified2) = 0 Then

	' Kill client first
	myLog.WriteLine logMessage + getDateTime() + "Log file was not change in " & sleepTime & " milis, KILL CLIENT"
	killProcess()
	myLog.WriteLine logMessage + getDateTime() + "Sleep after kill..."
	
	'
	' Run send error log
	'
	'myLog.WriteLine logMessage + getDateTime() + "RUN send error log!"
	'runSendErrorLog(rootDir)
	'myLog.WriteLine logMessage + getDateTime() + "Run SUCCESS!"
	
	WScript.Sleep 1000
	
	'
	' Run client after killing it.
	'
	myLog.WriteLine logMessage + getDateTime() + "RUN client again now!"
	runClient(rootDir)
	myLog.WriteLine logMessage + getDateTime() + "Run SUCCESS!"
	
	
Else
	myLog.WriteLine logMessage + getDateTime() + "Log file wasn't changed, App is running fine, OK"
End If

myLog.WriteLine logMessage + getDateTime() + "FINISH CRON..."
myLog.WriteLine logMessage + getDateTime() + "================================================================="
myLog.close()