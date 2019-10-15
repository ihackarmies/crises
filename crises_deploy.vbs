'**************************************************************
'						C.R.I.S.E.S.
' Compliance Reporting Information Search and Extraction System
'					Deployment Script	
'
'					Author: Alan Bairley
'
'						Version: 1.1
'				Date Revised: 22 May 2013
' 
' This script deploys the crises_interogation script to all 
' computer names / IP addresses contained within the file
' identifed by the computerListFile variable.

' Success/error results of the deployment are placed on the server 
' in the ./<computername> folder where '.' is the the directory
' of the crises_deployment script.
'
' If a computer already contains the crises_interrogation script,
' the existing script will be overwritten.
'
' Script Usage: crises_deploy.vbs
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
	serverScriptDir = "\\" & serverName & serverScriptPath

	Set NetworkObject = CreateObject("WScript.Network")

	Call writetoLog("\deploy_",".txt","====== " & now & " ==============================================")

	NetworkObject.MapNetworkDrive "", clientAdminShare, False', username, password

	clientScriptDir = clientAdminShare & clientScriptPath

	if not fso.FolderExists(clientScriptDir) then
		fso.CreateFolder(clientScriptDir)
		Call writetoLog("\deploy_",".txt","Creating folder " & clientScriptDir & " on " & ucase(ComputerName) & " ... DONE")
	else
		Call writetoLog("\deploy_",".txt","Folder " & clientScriptDir & " on " & ucase(ComputerName) & " already exists ... no action taken")
	end if

	fso.CopyFile serverScriptDir & agentScript, clientScriptDir & agentScript, True
	Call writetoLog("\deploy_",".txt","Creating file " & clientScriptDir & agentScript & " on " & ucase(ComputerName) & " ... DONE")

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
