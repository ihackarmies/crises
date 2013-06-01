'**************************************************************
'						C.R.I.S.E.S.
' Compliance Reporting Information Search and Extraction System
'					Scanning Script	
'
'					Author: Alan Bairley
'
'						Version: 1.1
'				Date Revised: 22 May 2013
' 
' This script invokes the crises_interrogation script on each computer
' listed in the file identified by the computerListFile varible.
'
' Success/error results of the scan are placed on the server in
' the ./<computername> folder where '.' is the the directory
' of the crises_scan script.
'
' Script Usage: crises_scan.vbs
'****************************************************************

Const windowsPath = "\windows\system32"
Const serverName = "fileserver1a"
Const serverScriptPath = "\officeshares\AOSC-SIG_(112th_SIG_BN)\crises_scripts"
Const serverScriptCmd = "\crises_interrogate.vbs"
Const computerListFile = "Computers.txt"

Dim args,logPath,computername,serverScript,username,password
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

	serverShare = "\\" & serverName & serverScriptPath
	serverScript = serverShare & serverScriptCmd

	Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objSWbemServices = objSWbemLocator.ConnectServer(computername,"root\cimv2")',username,password)
	Set objProcess = objSWbemServices.Get("Win32_Process")

	Set fso = CreateObject("Scripting.FileSystemObject")

	Call writetoLog("\scan_",".txt","====== " & now & " ==============================================")
			
	errReturn = objProcess.create("C:" & windowsPath & "\wscript.exe C:\agent_scripts\crises_interrogate.vbs", null, null, null)
	if errReturn = 0 then
		Call writetoLog("\scan_",".txt","Successful scan of computer: " & computername)
	else
		Call writetoLog("\scan_",".txt","Connection to " & computername & " could not be estblished due to error: " & errReturn)
	end if

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