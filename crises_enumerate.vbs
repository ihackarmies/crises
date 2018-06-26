'**************************************************************
'						C.R.I.S.E.S.
' Compliance Reporting Information Search and Extraction System
'					Enumeration Script	
'
'					Author: ihackcoinz
'
'						Version: 1.1
'				Date Revised: 22 May 2013
' 
' This script generates a file containing a listing of all
' Computer objects contained within an Active Directory OU.
' Default filename = "Computers.txt"
'
' Script Usage: crises_enumerate.vbs
'****************************************************************

Const listFile = "Computers.txt"
Const ADS_SCOPE_SUBTREE = 2  

Dim fso  
Set fso = CreateObject("Scripting.FileSystemObject")  
Dim outFile  
Set outFile = fso.OpenTextFile("Computers.txt", 2, True)  

Dim cn  
Set cn = CreateObject("ADODB.Connection")  
cn.Provider = "ADsDSOObject" 
cn.Open "Active Directory Provider" 
Dim cmd  
Set cmd = CreateObject("ADODB.Command")  
Set cmd.ActiveConnection = cn  
Set objRootDSE = GetObject("LDAP://RootDSE")
Dim ou  

strDNSDomain = objRootDSE.Get("DefaultNamingContext")
ou = "OU=child4,OU=child3,OU=Standard Workstations,OU=Domain Workstations"

cmd.CommandText = "SELECT name " & _  
                  "FROM 'LDAP://" & ou & "," & strDNSDomain &  "' " & _  
                  "WHERE objectClass='computer' " & _  
                  "ORDER BY name"    
cmd.Properties("Page Size") = 1000  
cmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE  

Dim  rs  
Set rs = cmd.Execute  
rs.MoveFirst  

Do Until rs.EOF  
'WScript.Echo rs(0)  
outFile.WriteLine rs(0)  
rs.MoveNext  
Loop 

outFile.Close  
Set outFile = Nothing 
Set fso = Nothing 
