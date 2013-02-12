'-------------------------------------------------------'
'              IM-One Global Installer                  '
'  Installation application for all IM-One Components.  '
'-------------------------------------------------------'
'  Written By: Christopher S. Bates     '
'  Written For: Forward Advantage Inc.  '
'---------------------------------------'
'  Version: 0.5.1           '
'  Date: February 12, 2013  '
'---------------------------'
'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'
'                                                                              Global Variables                                                                                                        '
'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'
Const sInstallerVersion = "0.5.1"

Const sComputer = "."												'Local computer
Const sBinDir = ".\Bin\"											'Directory where binary installers are stored
Const sRootDir = ".\"												'Sets the root directory for the installer
Const HKCR = &H80000000												'HKEY_CLASSES_ROOT
Const HKCU = &H80000001												'HKEY_CURRENT_USER
Const HKLM = &H80000002												'HKEY_LOCAL_MACHINE
Const HKUS = &H80000003												'HKEY_USERS
Const HKCC = &H80000005												'HKEY_CURRENT_CONFIG
Const sSupportNumber = "1(866) 636-9310"							'IM-One Support Phone Number
Const sSupportEmail = "im-onesupport@forwardadvantage.com"			'IM-One Support Email Address
Const sAppName = "IM-One Global Installer"							'IM-One Support Phone Number
Const sInstallLog = ".\InstallLog.txt"								'Location of the install log
Const ForAppending = 8												'Used for appending logs.


Dim sCheckRegHive, sCitrixRecNewVersion, sViewNewVersion, sAMNewVersion, sOneSignNewVersion, sIMRDPNewVersion ' Versions of MSI's contained in package and IM-OneInstaller Reg Hive - set in UpdateCheck()
Dim sWindowsArchitecture, sWindowsVersion ' Windows version and architecture - set in WinArch() and WinVer()
Dim sUserPath, sProgramFiles, sSoftwareRegistryHive ' Windows Variables based on version and architecture - set in WindowsCheck()

'----------------------------------------------------------------------------------------END OF GLOBAL----------------------------------------------------------------------------------------'
DelFile(sInstallLog)

sWindowsArchitecture = WinArch()

sWindowsVersion = WinVer()

WriteLog( _
Chr(13) &_
"******************************************************************" & Chr(13) &_
"*                     IM-One Installer Log                       *" & Chr(13) &_ 
"*  Installer Written By: Christopher S. Bates                    *" & Chr(13) &_
"*  Installer Written For: Forward Advantage Inc.                 *" & Chr(13) &_
"******************************************************************" & Chr(13) &_
Chr(13) & Chr(13) & Chr(13) & "!!!!-----------------------------------------Start Logging-----------------------------------------!!!!" & Chr(13))

UpdateCheck()

WriteLog(Chr(13) & "------------------------------Setting Global Variabls---------------------------------------" & Chr(13))

WindowsCheck()

WriteLog(Chr(13) & "------------------------------Check and Disable UAC---------------------------------------" & Chr(13))

DisableUAC()









'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'
'                                                                              Primary Functions                                                                                                       '
'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'


'----------------------------------------------------------------------------------------------------------------------------------------------------'
'                                         !!!Primary Function - Check For Updates!!!                                                                 '
'----------------------------------------------------------------------------------------------------------------------------------------------------'
'  Primary Function: Check for First Run. Then finds and compares the versions of the packeges currently installed vs. the ones in the new package.  '
'----------------------------------------------------------------------------------------------------------------------------------------------------'
'    If new packages are found update proceeds.  '
'    If no packages found, installer runs        '
'    If no updates found, installer exits        '
'------------------------------------------------'
Function UpdateCheck() ' Checks to see if this is the first time installing the package, or if there are any updates to be made.
	Dim sFirstRun
	Dim sCitrixRecCurrentVersion, sViewCurrentVersion, sAMCurrentVersion, sOneSignCurrentVersion, sIMRDPCurrentVersion
	
	'-----------------------------------------------------------------'
	'                             First Run                           '
	'  Checks for the IM-OneInstaller Reg Key, if not exist make it   '
	'-----------------------------------------------------------------'
	WriteLog(Chr(13) & "--------------------------------Checking First Run--------------------------------" & Chr(13))
	sCheckRegHive = "HKEY_LOCAL_MACHINE\SOFTWARE\IM-OneInstaller\"
	
	sFirstRun = KeyExists(HKLM, "SOFTWARE\IM-OneInstaller\")
	
	if sFirstRun = "False" Then
		
		WriteLog("Running as first time installation")
		WriteLog("Creating registry hive... " & sCheckRegHive)
		
		CreateRegKey(sCheckRegHive)
	
	Else

	WriteLog("Registry hive already found... " & sCheckRegHive)
		WriteLog("Running as an update")
	
	End if
	
	'-------------------------------------------------------------------------'
	'                    Check MSI File Versions                              '
	'  Checks OneSign and Authentication Manager MSI versions in the package  '
	'-------------------------------------------------------------------------'
	WriteLog(Chr(13) & "----------------------------Checking Packaged Versions-----------------------------------" & Chr(13))
	If "sWindowsArchitecture" = "x64" Then
		sOneSignNewVersion = GetMsiVersion(sBinDir & "OneSignAgentx64.msi")
		WriteLog("OneSign x64 MSI version... " & sOneSignNewVersion)
	
	Else 
		sOneSignNewVersion = GetMsiVersion(sBinDir & "OneSignAgent.msi")
		WriteLog("OneSign x86 MSI version... " & sOneSignNewVersion)
	End If
	
	sAMNewVersion = GetMsiVersion(sBinDir & "AMCLIENT.msi")
	WriteLog("Authentication Manager MSI version... " & sAMNewVersion)
	
	'-------------------------------------------------------------------------'
	'                    Check EXE File Versions                              '
	'  Checks XenApp, XenDesktop, and VMwareView EXE versions in the package  '
	'-------------------------------------------------------------------------'
	If "sWindowsArchitecture" = "x64" Then
		sViewNewVersion = GetExeVersion(sBinDir & "VMWare\ViewClientx64.exe")
		WriteLog("View Client x64 EXE version = " & sViewNewVersion)
	Else
		sViewNewVersion = GetExeVersion(sBinDir & "VMWare\ViewClient.exe")
		WriteLog("View Client x86 EXE version = " & sViewNewVersion)
	End If
	
	sCitrixRecNewVersion = GetExeVersion(sBinDir & "CitrixReceiver.exe")
	WriteLog("Citrix Receiver EXE version... " & sCitrixRecNewVersion)

	sIMRDPNewVersion = GetExeVersion(sBinDir & "IMRDP\IMONERDP.exe")
	WriteLog( "IM-RDP EXE Version = " & sIMRDPNewVersion)
	
	'----------------------------------------------------------'
	'               Check Current Installed Version            '
	'  Checks the currently installed version of all software  '
	'----------------------------------------------------------'
	WriteLog(Chr(13) & "------------------------------Checking Currently Installed Versions---------------------------------------" & Chr(13))
	
	sOneSignCurrentVersion = ReadRegValue(sCheckRegHive & "OneSignVer", 0)
	WriteLog("OneSign installed version... " & sOneSignCurrentVersion)
	
	sCitrixRecCurrentVersion = ReadRegValue(sCheckRegHive & "CTXVer", 0)
	WriteLog("Citrix Man installed Version... " & sCitrixRecCurrentVersion)
	
	sViewCurrentVersion	= ReadRegValue(sCheckRegHive & "VMWareVer", 0)
	WriteLog("VMWare installed Version... " & sViewCurrentVersion)
	
	sAMCurrentVersion = ReadRegValue(sCheckRegHive & "AuthManVer", 0)
	WriteLog("Auth Man installed Version... " & sAMCurrentVersion)
	
	sIMRDPCurrentVersion = ReadRegValue(sCheckRegHive & "IM-RDPVer", 0)
	WriteLog("IMRDP installed Version... " & sIMRDPCurrentVersion)
	
	' Check for Updates based on registry keys if none are found exit.
	If sOneSignNewVersion <> sOneSignCurrentVersion _
	OR sCitrixRecNewVersion <> sCitrixRecCurrentVersion _
	OR sViewCurrentVersion <> sViewNewVersion _
	OR sAMCurrentVersion <> sAMNewVersion _
	OR sIMRDPCurrentVersion <> sIMRDPNewVersion _
	Then
		WriteLog ("Update Found, proceeding with installation")
	Else	

		WriteLog("No updates found")
		WriteLog(" Exiting installer")
		
		MsgBox "There are no updates at this time."
		Wscript.Quit
	End If
	
End Function

'-------------------------------------------------------------------------------------------------------------------------'
'                                         !!!Primary Function - Gather Windows Information!!!                             '
'-------------------------------------------------------------------------------------------------------------------------'
'  Primary Function: Checks to see what version and architecure of Windows you are running and determines compatibility.  '
'-------------------------------------------------------------------------------------------------------------------------'
'  If Running Windows XP Check x86 or x64         '
'        If x86 set x86  XP variables             '
'        If x64 set x64 XP variables              '
'  If Running Windows Vista Check for x86 or x64  '
'        If x86 set x86 Vista variables           '
'        If x64 set x64 Vista variables           '
'  If Running Windows 7 Check for x86 or x64      '
'        If x86 set x86 7 variables               '
'        If x64 set x64 7 variables               '
'  Else Exit Incompaitible                        '
'-------------------------------------------------'
Function WindowsCheck()

	' Checks for the version of Windows and then calls function to set global variables.
	If sWindowsVersion = 2600 Then
		WriteLog("Found Windows XP")
		SetXPVariables()
		
	ElseIf sWindowsVersion = 6000 _
	OR sWindowsVersion = 6002 _
	OR sWindowsVersion = 7600 _
	OR sWindowsVersion = 7601 _
	Then
		
		WriteLog("Found Windows 7 or Windows Vista")
		SetWin7Variables()
	
	Else
	
		WriteLog("Unsupported version of Windows: " & sWindowsVersion) 'Log
		ErrorBox("You are running an unsupported version of Windows.")

		
	End If
	
	WriteLog(Chr(13) &_
			"	User path set to... " & sUserPath & Chr(13) &_
			"	Program files path set to... " & sProgramFiles & Chr(13) &_
			"	Software registry hive set to ... " & sSoftwareRegistryHive & Chr(13))
			
End Function
'----------------------------------------------------------------------------------------------------'
'                        !!!Primary Function - Disable User Access Control!!!                        '
'----------------------------------------------------------------------------------------------------'
'  Primary Function: Checks for Windows 7, and then prompts user to see if they wish to disable UAC  '
'----------------------------------------------------------------------------------------------------'
' If Windows 7                    '
'     Prompt User to disable UAC  '
' Else                            '
'    Skip                         '
'---------------------------------'

Function DisableUAC()


	Dim sResponse, sUACRegValue

	sUACRegValue = ReadRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA", 1)

	If sWindowsVersion = 6000 _
	OR sWindowsVersion = 6002 _
	OR sWindowsVersion = 7600 _
	OR sWindowsVersion = 7601 _
	AND sUACRegValue <> 0 Then

		sResponse = MsgBox("Would you like to disable the UAC?", 4, sAppName) 'User selects to disable UAC or not

		If sResponse = 6 Then
			WriteLog("Disabling UAC")
			SetRegDword HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA", 0
		Else
			WriteLog("User Action: Do not disable UAC")
		End If
	
	Else
		WriteLog("UAC is disabled or not present")
	
	End If

	sUACRegValue = ReadRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA", 1) 'Logging for UAc
	WriteLog("UAC Value set to... " & sUACRegValue)	

End Function



'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'
'                                                                              Secondary Functions                                                                                                     '
'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'
'--------------------------------------------------'
'  Function: Sets Variables if Windows 7 is Found  '
'--------------------------------------------------'
' Used in WindowsCheck()
Function SetXPVariables()
		sUserPath = "C:\Documents and Settings\"
		sProgramFiles = "C:\ProgramFiles\"
		sSoftwareRegistryHive = "SOFTWARE\"
End Function


'--------------------------------------------------'
'  Function: Sets Variables if Windows 7 is Found  '
'--------------------------------------------------''
' Used in WindowsCheck()
Function SetWin7Variables()
	If  sWindowsArchitecture = "x86" Then
	
		WriteLog("Setting Variables for Windows 7x86") 'Log
		
		sUserPath = "C:\Users\"
		sProgramFiles = "C:\ProgramFiles\"
		sSoftwareRegistryHive = "SOFTWARE\"
		
		
	ElseIf sWindowsArchitecture = "x64" Then
	
		WriteLog("Setting Variables for Windows 7x64") 'Log
		sUserPath = "C:\Users\"
		sProgramFiles = "C:\ProgramFiles (x86)\"
		sSoftwareRegistryHive = "SOFTWARE\Wow6432Node\"
		
	Else
		
		WriteLog("Could not determine Operating System Architecture") 'Log
		
		ErrorBox("Could not determine Operating System Architecture")
	End If
		
End Function

'-----------------------------------------'
'  Function: Get Windows OS Architecture  '
'-----------------------------------------'
' Used in WindowsCheck()
Function WinArch()
	Dim wshShell, osType
	
	Set wshShell = CreateObject("WScript.Shell")
	OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
	
	WriteLog("Determining Windows Architecture") 'Log
	
	If OsType = "x86" then
		WinArch = "x86"
	elseif OsType = "AMD64" then
		WinArch = "x64"
	end if
	
	WriteLog("Windows Architecture... " & WinArch) 'Log
	
End Function

'------------------------------------------------------'
'  Function: Get Windows Version baed on Build Number  '
'------------------------------------------------------'
' Used in WindowsCheck()
Function WinVer()
	Dim objWMIService, oss, os, dtmConvertedDate
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
	Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	
	WriteLog("Determining Windows Version Number") 'Log
	
	For Each os in oss
		WinVer =  os.BuildNumber
	Next
	
	WriteLog("Windows Version... " & WinVer) 'Log
End Function

'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'
'                                                                               Minor Functions                                                                                                        '
'!!!!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------!!!!'
'--------------------------------------------------------------------*'
'                    Check If Registry Key Exists                     '
'  You can pass a registry key to this function to see if it exists.  '
'--------------------------------------------------------------------*'
Function KeyExists(hDefKey, strKeyPath)

	Dim oReg: Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\default:StdRegProv")

	If oReg.EnumKey(hDefKey, strKeyPath) = 0 Then
		KeyExists = True
	Else
		KeyExists = False
	End If

End Function

'----------------------------------------------------------------------------'
'                          Check MSI version                                 '
'  You can pass an MSI file to this function to get the version of the MSI.  '
'----------------------------------------------------------------------------'
Function GetMsiVersion(sMSIFile)
	Dim  Installer, Database, SQL, View, Record
	
	Set Installer = CreateObject("WindowsInstaller.Installer")
	
	Set Database = Installer.OpenDatabase(sMSIFile, 0)
	
	SQL = "SELECT * FROM Property WHERE Property = 'ProductVersion'"
	
	Set View = DataBase.OpenView(SQL)
	
	View.Execute
	
	Set Record = View.Fetch
	
	GetMsiVersion = Record.StringData(2)

End Function


'----------------------------------------------------------------------------'
'                            Check EXE version                               '
'  You can pass an EXE file to this function to get the version of the EXE.  '
'----------------------------------------------------------------------------'
Function GetExeVersion(sExeVersion)
	
	'MsgBox sExeVersion
	GetExeVersion = CreateObject("Scripting.FileSystemObject").GetFileVersion(sExeVersion)

End Function


'------------------------------------------------------------------'
'                     Get Registry Key Value                       '
'  Passing a registry key to this function will return its value.  '
'------------------------------------------------------------------'
Function ReadRegValue(sRegKey, sDefault)
	Dim wshShell, value

	On Error Resume Next
	Set wshShell = CreateObject("WScript.Shell")
	value = wshShell.RegRead(sRegKey)

	if err.number <> 0 then
		'WriteLog("Key not found... " & sRegKey)
		ReadRegValue= sDefault
	else
		'WriteLog("Found Value... " & sRegKey)
		ReadRegValue = wshShell.RegRead(sRegKey)
	end if
	
	set wshShell = nothing
	
End Function

'----------------------------------------------------------'
'                     Create Registry Key                  '
'  Imput full registry key path and this will generate it  '
'----------------------------------------------------------'

Function CreateRegKey(sKey)
	Dim wshShell
	Set wshShell = CreateObject( "WScript.Shell" )
	
	wshShell.RegWrite sKey, ""
	WriteLog("Created Registry Key... " & sKey)
	
	Set wshShell = nothing

End Function


'---------------------------'
'         Write Log         '
'  Writes line to logfile.  '
'---------------------------'
Sub WriteLog(sLogLine)
	Dim oFileSystem, oFile
	set oFileSystem = CreateObject("Scripting.FileSystemObject")
	set oFile = oFileSystem.OpenTextFile(sInstallLog, ForAppending, True)

	oFile.WriteLine(Now & "  |  " & sLogLine)

	oFile.Close

End Sub

'--------------------------'
'       Delete File        '
'  Deletes specified file  '
'--------------------------'
Sub DelFile(sFileName)
	dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject")
	
	filesys.CreateTextFile sFileName, True
	
	If filesys.FileExists(sFileName) Then
	   filesys.DeleteFile sFileName
	   WriteLog("Deleted... " & sFileName)
	End If 

End Sub