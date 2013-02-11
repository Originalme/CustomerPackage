Option Explicit
'*************************************************************************************'
'                                  Update Check                                       '
'  Checks to see if there are updates available for any of the following components.  '
'  Written By: Christopher S. Bates                                                   '
'*************************************************************************************'
'  Version 0.9.2  '
'*****************'

Const sComputer = "."
Const sBinDir = "..\Bin\"
Const sRootDir = ".\"
Const HKCR = &H80000000 'HKEY_CLASSES_ROOT
Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const HKUS = &H80000003 'HKEY_USERS
Const HKCC = &H80000005 'HKEY_CURRENT_CONFIG
Const sSupportNumber = "1(866) 636-9310"
Const sSupportEmail = "im-onesupport@forwardadvantage.com"
Const sAppName = "IM-One Global Installer"

'--------------------------------------------------------------------END OF GLOBAL--------------------------------------------------------------------'
UpdateCheck()

Function UpdateCheck()
	Dim sCheckRegHive, sFirstRun
	Dim sCitrixRecCurrentVersion, sCitrixRecNewVersion
	Dim sViewCurrentVersion, sViewNewVersion
	Dim sAMCurrentVersion, sAMNewVersion
	Dim sOneSignCurrentVersion, sOneSignNewVersion
	Dim sIMRDPCurrentVersion, sIMRDPNewVersion
	
	'*****************************************************************'
	'                             First Run                           '
	'  Checks for the IM-OneInstaller Reg Key, if not exist make it   '
	'*****************************************************************'

	sCheckRegHive = "HKEY_LOCAL_MACHINE\SOFTWARE\IM-OneInstaller\"
	
	
	sFirstRun = KeyExists(HKLM, "SOFTWARE\IM-OneInstaller\")
	
	if sFirstRun = "False" Then
		CreateRegKey(sCheckRegHive)
	End if
	
	
	'**********************************************************'
	'                    Check MSI File Versions               '
	'  Checks OneSign and Authentication Manager MSI Versions  '
	'**********************************************************'
	
	If "x64" = "x64" Then
		sOneSignNewVersion = GetMsiVersion(sBinDir & "OneSignAgentx64.msi")
		'MsgBox "OneSign x64 version = " & sOneSignNewVersion
	
	Else 
		sOneSignNewVersion = GetMsiVersion(sBinDir & "OneSignAgent.msi")
		'MsgBox "OneSign x86 version = " & sOneSignNewVersion
	End If
	
	sAMNewVersion = GetMsiVersion(sBinDir & "AMCLIENT.msi")
	MsgBox "AM NEW VERSION = " & sAMNewVersion
	
	'**********************************************************'
	'                    Check EXE File Versions               '
	'  Checks XenApp, XenDesktop, and VMwareView MSI Versions  '
	'**********************************************************'
	
	sAMNewVersion = GetMsiVersion(sBinDir & "AMCLIENT.msi")
	'MsgBox "Authentication Manager version = " & sAMNewVersion

	
	sCitrixRecNewVersion = GetExeVersion(sBinDir & "CitrixReceiver_3_4.exe")
	'MsgBox "Citrix Receiver Version = " & sCitrixRecNewVersion

	If "x64" = "x64" Then
		sViewNewVersion = GetExeVersion(sBinDir & "VMWare\ViewClientx64.exe")
		MsgBox "View Client Version = " & sViewNewVersion
	Else
		sViewNewVersion = GetExeVersion(sBinDir & "VMWare\ViewClient.exe")
		MsgBox "View Client Version = " & sViewNewVersion
	End If
	
	'sAMNewVersion = GetExeVersion(sBinDir & "AMCLIENT.exe")
	'MsgBox "Authentication Manager Client Version = " & sAMNewVersion

	sIMRDPNewVersion = GetExeVersion(sBinDir & "IMRDP\IMONERDP.exe")
	'MsgBox "RDP Version = " & sIMRDPNewVersion

	'*********************************************************'
	'               Check Current Installed Version           '
	'  Checks the currently installed version of all software '
	'*********************************************************'
	
	sOneSignCurrentVersion = ReadRegValue(sCheckRegHive & "OneSignVer", 0)
	'MsgBox "OS Current Version = " & sOneSignCurrentVersion
	
	sCitrixRecCurrentVersion = ReadRegValue(sCheckRegHive & "CTXVer", 0)
	'MsgBox "Citrix Man Current Version = " & sCitrixRecCurrentVersion
	
	sViewCurrentVersion	= ReadRegValue(sCheckRegHive & "VMWareVer", 0)
	'MsgBox "VMWare Current Version = " & sViewCurrentVersion
	
	sAMCurrentVersion = ReadRegValue(sCheckRegHive & "AuthManVer", 0)
	'MsgBox "Auth Man Current Version = " & sAMCurrentVersion
	
	sIMRDPCurrentVersion = ReadRegValue(sCheckRegHive & "IM-RDPVer", 0)
	'MsgBox "IMRDP Current Version = " & sIMRDPCurrentVersion
	

	
	
	If sOneSignNewVersion <> sOneSignCurrentVersion Then
		main()
	End If	
	
	If sCitrixRecNewVersion <> sCitrixRecCurrentVersion Then
		main()
	End If
	
	MsgBox "There are no updates at this time."
	Wscript.Quit
	
	
End Function

Function main()
	MsgBox "Made it to Main"
	Wscript.Quit
End Function

'***********************************************************************************************************************************************************************************************************************'
'*****************************************************************************************************SUB FUNCTIONS*****************************************************************************************************'
'***********************************************************************************************************************************************************************************************************************'

'****************************************************************************'
'                          Check MSI version                                 '
'  You can pass an MSI file to this function to get the version of the MSI.  '
'****************************************************************************'
Function GetMsiVersion(sMSIFile)
	Dim  Installer, Database, SQL, View, Record
	
	Set Installer = CreateObject("WindowsInstaller.Installer")
	
	Set Database = Installer.OpenDatabase(sMSIFile, 0)
	
	SQL = "SELECT * FROM Property WHERE Property = 'ProductVersion'"
	
	Set View = DataBase.OpenView(SQL)
	
	View.Execute
	
	Set Record = View.Fetch
	
	GetMsiVersion = Record.StringData(2)
	'MsgBox "MSI File Version = " & sMSIFile

End Function

'****************************************************************************'
'                          Check EXE version                                 '
'  You can pass an EXE file to this function to get the version of the EXE.  '
'****************************************************************************'
Function GetExeVersion(sExeVersion)
	
	'MsgBox sExeVersion
	GetExeVersion = CreateObject("Scripting.FileSystemObject").GetFileVersion(sExeVersion)

End Function	

'*********************************************************************'
'                       Check If Reg Key Exists                       '
'  You can pass a registry key to this function to see if it exists.  '
'*********************************************************************'
Function KeyExists(hDefKey, strKeyPath)

	Dim oReg: Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\default:StdRegProv")

	If oReg.EnumKey(hDefKey, strKeyPath) = 0 Then
		KeyExists = True
	Else
		KeyExists = False
	End If

End Function

'******************************************************************'
'                        Get Registry Key Value                    '
'  Passing a registry key to this function will return its value.  '
'******************************************************************'
Function ReadRegValue(sRegKey, sDefault)
	Dim wshShell, value

	On Error Resume Next
    Set wshShell = CreateObject("WScript.Shell")
    value = wshShell.RegRead(sRegKey)

    if err.number <> 0 then
       ReadRegValue= sDefault
    else
        ReadRegValue = wshShell.RegRead(sRegKey)
    end if
	
    set wshShell = nothing
	
End Function

'**********************************************************'
'                     Create Registry Key                  '
'  Imput full registry key path and this will generate it  '
'**********************************************************'

Function CreateRegKey(sKey)
	Dim wshShell
	Set wshShell = CreateObject( "WScript.Shell" )
	
	wshShell.RegWrite sKey, ""
	
	Set wshShell = nothing

End Function
	