'File Name: Logon.vbs
'Version: v1.0, 11/25/2022
'Author: Justin Grimes, 11/25/2022

' --------------------------------------------------
'Declare explicit variables to be used during the session.
Option Explicit 
Dim fileSystem, windowsVersion, oShell, computerName, message, logPath, strSafeDate, strSafeTime, _
 strDateTime, logfile, objLogFile, appPath, verbose, email, logging, logData, dataDir, computerDir, _
 OutputBox, processArch, sysArch, messageData, archType, checkupResults, error, dxDiagInfo, _
 companyName, companyAbbr, userName, strSafeTimeA, strSafeDateA
' --------------------------------------------------

' --------------------------------------------------
'Set variables values for the session.
Set oShell = WScript.CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject") 
userName = CreateObject("WScript.Network").UserName
computerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strSafeDate = DatePart("yyyy",Date)&Right("0"&DatePart("m",Date), 2)&Right("0"&DatePart("d",Date), 2)
strSafeDateA = DatePart("yyyy",Date)&"-"&Right("0"&DatePart("m",Date), 2)&"-"&Right("0"&DatePart("d",Date), 2)
strSafeTime = Right("0"&Hour(Now), 2)&Right("0"&Minute(Now), 2)&Right("0"&Second(Now), 2)
strSafeTimeA = Right("0"&Hour(Now), 2)&":"&Right("0"&Minute(Now), 2)&":"&Right("0"&Second(Now), 2)
strDateTime = strSafeDate&"-"&strSafeTime
  ' ----------
  ' Company Specific variables.
  ' Change the following variables to match the details of your organization.
  
  ' The "appPath" is the full absolute path for the script directory, with trailing slash.
  appPath = "\\server\Scripts\Logon\"
  ' The "logPath" is the full absolute path for where network-wide logs are stored.
  logPath = "\\server\Logs"
  ' The "companyName" the the full, unabbreviated name of your organization.
  companyName = "The Company Inc."
  ' The "companyAbbr" is the abbreviated name of your organization.
  companyAbbr = "TCI"
  ' ----------
logfile = logPath&"\"&computerName&"-"&userName&"-"&strDateTime&"-logon.txt"
message = "The user "&userName&" has logged into the system "&computerName&" on "&strSafeDateA&" at "&strSafeTimeA&"."
' --------------------------------------------------

' --------------------------------------------------
'A function to create a log file when -l is set.
'Returns "True" if logfile exists, "False" on error.
Function CreateLog(logfile, message)
  error = True
  If message <> "" Then
    Set objLogfile = fileSystem.CreateTextFile(logfile, True)
    objLogfile.WriteLine(message)
    objLogfile.Close
  End If
  If fileSystem.FileExists(logfile) Then
    error = False
  End If
  CreateLog = error
End Function
' --------------------------------------------------

' --------------------------------------------------
' The main logic of the program which makes use of the functions above.
CreateLog logfile, message
' --------------------------------------------------