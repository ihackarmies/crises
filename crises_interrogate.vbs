'**************************************************************
'						C.R.I.S.E.S.
' Compliance Reporting Information Search and Extraction System
'					Interrogation Script	
'
'					Author: Alan Bairley
'
'						Version: 1.1
'				Date Revised: 22 May 2013
'
' This script interrogates a computer system and records the
' results of each interrogation in a separate file:
'
' Installed Software Report = software_<computername>
' Running Processes Report = processes_<computername>
' System Information Report = system_<computername>
'
' All reports are placed in the .\<computername> folder,
' where '.' is the location where the crises_interrogate 
' script is located.
'
' An optional 'computername' parameter may be passed to the
' script in order to target a remote machine for interrogation.
' By default the computer targeted is the localhost.
'
' Script Usage: crises_interrogate.vbs [/computer:<computername>]
'****************************************************************

Dim args,logPath,computername
Const logHistory = 2
Set args = Wscript.Arguments.Named

computername = args.Item("computer")

if computername = "" then
	computername = getComputer()
end if

logPath = getLogPath()

Call getProcessInfo

Call getSoftwareInfo

Call getSystemInfo

Function getProcessInfo
On Error Resume Next
Dim objProcess,process,strNameOfUser
Call writetoLog("\processes_",".csv","====== " & now & " ==============================================")
Call writetoLog("\processes_",".csv","Name,Executable Path,Priority,Session Id,Owner,Handle Count,Thread Count")
Set objProcess = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & computername & "\root\cimv2").ExecQuery _
				 ("Select * From Win32_Process")
if not (errorChecking (computername)) then
	for each process in objProcess
		if process.name <> "System Idle Process" and process.name <> "System" then
			Return = process.GetOwner(strNameOfUser)
			Call writetoLog("\processes_",".csv",process.name & "," & process.executablepath & "," & process.priority & "," & process.sessionid & "," & strNameOfUser & "," & process.handlecount & "," & process.threadcount)
		end if
	next
end if
Set objProcess = nothing
End Function

Function getSoftwareInfo
On Error Resume Next
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & computername & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_Product")
Call writetoLog("\software_",".csv","====== " & now & " ==============================================")
Call writeToLog("\software_",".csv", _
"Description" & "," & "Identifying Number" & "," & _
"Install Date" & "," & "Install Location" & "," & _
"Install State" & "," & "Name" & "," & _
"Version" & "," _
 & "Vendor")
if not (errorChecking(computername)) then
	for each objSoftware in colSoftware
	 Call writeToLog("software_",".csv", _
	 objSoftware.Description & "," & _
	 objSoftware.IdentifyingNumber & "," & _
	 objSoftware.InstallDate & "," & _
	 objSoftware.InstallLocation & "," & _
	 objSoftware.InstallState & "," & _
	 objSoftware.Name & "," & _
	 objSoftware.Version & "," & _
	 objSoftware.Vendor)
	next
end if
Set objWMIService = nothing
End Function

Function getSystemInfo
On Error Resume Next
Call writetoLog("\system_",".txt","====== " & now & " ==============================================")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & computername & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
if not (errorChecking(computername)) then
set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
for each objComputer in colSettings 
    Call writetoLog("\system_",".txt","System Name: " & objComputer.Name)
    Call writetoLog("\system_",".txt","Logged on User: " & objComputer.UserName)
    Call writetoLog("\system_",".txt","System Manufacturer: " & objComputer.Manufacturer)
    Call writetoLog("\system_",".txt","System Model: " & objComputer.Model)
    Call writetoLog("\system_",".txt","Time Zone: " & objComputer.CurrentTimeZone)
    Call writetoLog("\system_",".txt","Total Physical Memory: " & _
        objComputer.TotalPhysicalMemory)
next
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
for each objOperatingSystem in colSettings 
    Call writetoLog("\system_",".txt","OS Name: " & objOperatingSystem.Name)
    Call writetoLog("\system_",".txt","Version: " & objOperatingSystem.Version)
    Call writetoLog("\system_",".txt","Service Pack: " & _
        objOperatingSystem.ServicePackMajorVersion _
            & "." & objOperatingSystem.ServicePackMinorVersion)
    Call writetoLog("\system_",".txt","OS Manufacturer: " & objOperatingSystem.Manufacturer)
    Call writetoLog("\system_",".txt","Windows Directory: " & _
        objOperatingSystem.WindowsDirectory)
    Call writetoLog("\system_",".txt","Locale: " & objOperatingSystem.Locale)
    Call writetoLog("\system_",".txt","Available Physical Memory: " & _
        objOperatingSystem.FreePhysicalMemory)
    Call writetoLog("\system_",".txt","Total Virtual Memory: " & _
        objOperatingSystem.TotalVirtualMemorySize)
    Call writetoLog("\system_",".txt","Available Virtual Memory: " & _
        objOperatingSystem.FreeVirtualMemory)
next
set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_Processor")
for each objProcessor in colSettings 
    Call writetoLog("\system_",".txt","Processor Arch Type: " & objProcessor.Architecture)
    Call writetoLog("\system_",".txt","Processor: " & objProcessor.Description)
next
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_BIOS")
for each objBIOS in colSettings 
    Call writetoLog("\system_",".txt","BIOS Version: " & objBIOS.Version)
	Call writetoLog("\system_",".txt","BIOS Serial Number: " & objBIOS.SerialNumber)
next
end if
set objWMIService = nothing
End Function

Function writetoLog(prefix,suffix,comment)
Dim fso,objTextStream
Set fso = CreateObject ("Scripting.FileSystemObject")
if fso.FileExists(logPath & prefix & computername & suffix) then
	set objFile = fso.GetFile(logPath & prefix & computername & suffix)
	if DateDiff("s",objFile.DateLastModified,now) > 90 then
		for i=(logHistory-1) to 1 Step -1
			if fso.FileExists(logPath & prefix & computername & suffix & ".old" & i) then
				fso.CopyFile logPath & prefix & computername & suffix & ".old" & i , logPath & prefix & computername & suffix & ".old" & (i+1), True
			end if
		next
		fso.CopyFile logPath & prefix & computername & suffix , logPath & prefix & computername & suffix & ".old1", True
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

Function getComputer()
	Dim objNet
	Set objNet = WScript.CreateObject("WScript.Network") 
	getComputer = objNet.ComputerName 
	Set objNet = Nothing 
End Function

Function errorChecking(ComputerName) 
errorChecking = False 
if err.number <> 0 then 
	Call writetoLog("\error_",".txt","Unable to connect to " & ucase(ComputerName))
	err.Clear () 
	errorChecking = True
end if 
end Function