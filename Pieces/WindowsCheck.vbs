Option Explicit
'-------------------------------------------------------------------------------------'
'   Windows Architecture                                                              '
'  Script checks the architecture recognized by a Windows operating system            '
'  Written By: Christopher S. Bates   '
'-------------------------------------------------------------------------------------'
'  Version 0.5.2  '
'-----------------'

Const sComputer = "."												'Local computer
Const sBinDir = "..\Bin\"											'Directory where binary installers are stored
Const sRootDir = ".\"												'Sets the root directory for the installer
Const HKCR = &H80000000												'HKEY_CLASSES_ROOT
Const HKCU = &H80000001												'HKEY_CURRENT_USER
Const HKLM = &H80000002												'HKEY_LOCAL_MACHINE
Const HKUS = &H80000003												'HKEY_USERS
Const HKCC = &H80000005												'HKEY_CURRENT_CONFIG
Const sSupportNumber = "1(866) 636-9310"							'IM-One Support Phone Number
Const sSupportEmail = "im-onesupport@forwardadvantage.com"			'IM-One Support Email Address
Const sAppName = "IM-One Global Installer"							'IM-One Support Phone Number

'--------------------------------------------------------------------END OF GLOBAL--------------------------------------------------------------------'
Dim sWindowsArchitecture, sWindowsVersion


sWindowsArchitecture = WinArch()
sWindowsVersion = WinVer()


' Checks for the version of Windows and then calls function to set global variables.
If sWindowsVersion = 2600 Then
	SetXPVariables()
ElseIf sWindowsVersion = 6000 _
OR sWindowsVersion = 6002 _
OR sWindowsVersion = 7600 _
OR sWindowsVersion = 7601 _
Then
	SetWin7Variables()
Else
	ErrorBox("You are running an unsupported version of Windows.")

End If

'---------------------------------------------------------FUNCTIONS---------------------------------------------------------'

'------------------------------------------------------'
'  Function: Get Windows Version baed on Build Number  '
'------------------------------------------------------'
Function WinVer()
	Dim objWMIService, oss, os, dtmConvertedDate
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
	Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	
	For Each os in oss
		WinVer =  os.BuildNumber
	Next

End Function

'-----------------------------------------'
'  Function: Get Windows OS Architecture  '
'-----------------------------------------'
Function WinArch()
	Dim wshShell, osType
	
	Set wshShell = CreateObject("WScript.Shell")
	OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
	
	If OsType = "x86" then
		WinArch = "x86"
	elseif OsType = "AMD64" then
		WinArch = "x64"
	end if
	
End Function

'--------------------------------------------------'
'  Function: Sets Variables if Windows 7 is Found  '
'--------------------------------------------------'
Function SetWin7Variables()
	If  sWindowsArchitecture = "x86" Then
		MsgBox "Windows 7 x86"
	ElseIf sWindowsArchitecture = "x64" Then
		MsgBox "Windows 7 x64"
	Else
		ErrorBox("Could not determine Operating System Architecture")
	End If
		
End Function

'---------------------------------------------------'
'  Function: Sets Variables if Windows XP is Found  '
'---------------------------------------------------'
Function SetXPVariables()
	If  sWindowsArchitecture = "x86" Then
		MsgBox "Windows XP x86"
	ElseIf sWindowsArchitecture = "x64" Then
		MsgBox "Windows XP x64"
	Else
		ErrorBox("Could not determine Operating System Architecture")
	End If
End Function



'---------------------------------------------------------'
'  Function: Displays error message with sMsg as content  '
'---------------------------------------------------------'
Function ErrorBox(sMsg)

	ErrorBox = MsgBox ( _
	"ERROR: " & vbNewLine &_
	sMsg & vbNewLine & vbNewLine &_
	"-------------------------------------------------------------" & vbNewLine & vbNewLine &_
	"Please Contact Forward Advantage." & vbNewLine &_
	"Phone: " & sSupportNumber & vbNewLine &_
	"Email: " & sSupportEmail, _
	vbCritical,sAppName _
	)
End Function


'---------------------------------------------------------NOTES BELOW PLEASE IGNORE---------------------------------------------------------'


Function SetXPVariables()
	MsgBox "Windows XP Variables Being Set."
End Function

'Reminder of all the windows info commands
Function Notes()
	Dim objWMIService, oss, os, dtmConvertedDate
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
	Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	
	For Each os in oss
		Wscript.Echo "Boot Device: " & os.BootDevice
		Wscript.Echo "Build Number: " & os.BuildNumber
		Wscript.Echo "Build Type: " & os.BuildType
		Wscript.Echo "Caption: " & os.Caption
		Wscript.Echo "Code Set: " & os.CodeSet
		Wscript.Echo "Country Code: " & os.CountryCode
		Wscript.Echo "Debug: " & os.Debug
		Wscript.Echo "Encryption Level: " & os.EncryptionLevel
		Wscript.Echo "Licensed Users: " & os.NumberOfLicensedUsers
		Wscript.Echo "Organization: " & os.Organization
		Wscript.Echo "OS Language: " & os.OSLanguage
		Wscript.Echo "OS Product Suite: " & os.OSProductSuite
		Wscript.Echo "OS Type: " & os.OSType
		Wscript.Echo "Primary: " & os.Primary
		Wscript.Echo "Registered User: " & os.RegisteredUser
		Wscript.Echo "Serial Number: " & os.SerialNumber
		Wscript.Echo "Version: " & os.Version
	Next
	
End Function

