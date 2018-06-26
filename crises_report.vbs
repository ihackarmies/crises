'**************************************************************
'						C.R.I.S.E.S.
' Compliance Reporting Information Search and Extraction System
'					Reporting Script	
'
'					Author: ihackcoinz
'
'						Version: 1.1
'				Date Revised: 22 May 2013
'
' This script pulls the interrogation reports from each computer
' listed in the file identified by the computerListFile variable
' and places the reports on the server in the folder: ./<computername>
' where '.' is the directory of the crises_report script.
'
' Script Usage: crises_report.vbs
'****************************************************************

Const windowsPath = "\windows\system32"
Const clientScriptPath = "\agent_scripts"
Const serverName = "fileserver1a"
Const serverScriptPath = "\officeshares\temp\crises_scripts"
Const agentScript = "\crises_interrogate.vbs"
Const computerListFile = "Computers.txt"

Dim args,logPath,computername,clientScriptDir,serverScriptDir,username,password
Set args = Wscript.Arguments.Named

'computername = args.Item("computer")
'username = args.Item("user")
'password = args.Item("password")

'if computername = "" then
'	computername = getComputer()
'end if

Set fso = CreateObject("Scripting.FileSystemObject")
Set listFile = fso.OpenTextFile(computerListFile)

do while not listFile.AtEndOfStream 
	
	On Error Resume Next

	computername =  listFile.ReadLine()

	logPath = getLogPath()

	clientAdminShare = "\\" & computername & "\C$"
	serverScriptDir = "\\" & serverName & serverScriptPath & "\" & computername

	Set NetworkObject = CreateObject("WScript.Network")
	Set fso = CreateObject("Scripting.FileSystemObject")

	Call writetoLog("\report_",".txt","====== " & now & " ==============================================")

	NetworkObject.MapNetworkDrive "", clientAdminShare, False', username, password

	clientScriptDir = clientAdminShare & clientScriptPath & "\" & computername

	if not fso.FolderExists(clientScriptDir) then
		Call writetoLog("\report_",".txt","Folder " & clientScriptDir & " on " & ucase(computername) & " does not exist... exiting")
	else
		fso.CopyFolder clientScriptDir, serverScriptDir, True
		Call writetoLog("\report_",".txt","Copied all files from " & clientScriptDir & " to " & serverScriptDir)
	end if

	NetworkObject.RemoveNetworkDrive clientAdminShare, True, False
	Set NetworkObject = Nothing

	if Err then
		Call writetoLog("\error_",".txt","====== " & now & " ==============================================")
		Call writetoLog("\error_",".txt","Error '" & Err.Number & "' encountered.  Unable to connect to computer: " & computername)
	end if
	
loop

Function getComputer()
	Dim objNet
	Set objNet = WScript.CreateObject("WScript.Network") 
	getComputer = objNet.ComputerName 
	Set objNet = Nothing 
End Function

Function writetoLog(prefix,suffix,comment)
Dim fso,objTextStream
Set fso = CreateObject ("Scripting.FileSystemObject")
if fso.FileExists(logPath & prefix & computername & suffix) then
	set objFile = fso.GetFile(logPath & prefix & computername & suffix)
	if DateDiff("s",objFile.DateLastModified,now) > 30 then
		fso.DeleteFile(logPath & prefix & computername & suffix)
	end if
end if
Set objTextStream = fso.OpenTextFile(logPath & prefix & computername & suffix, 8,True)
	objTextStream.WriteLine(comment)
	objTextStream.close
Set objTextStream = nothing
Set fso = nothing
End Function

Function getLogPath()
Dim temp,temp2,fso
Set fso = CreateObject ("Scripting.FileSystemObject")
temp = split(wscript.scriptfullname,"\")
for i = 0 to ubound(temp) - 1
	temp2 = temp2 & temp(i) & "\"
next
temp2 = temp2 & computername & "\"
if not fso.FolderExists(temp2) then
	fso.CreateFolder(temp2)
end if
getLogPath = temp2
Set fso = nothing
End Function
