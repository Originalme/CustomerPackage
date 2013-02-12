Dim sWindowsArchitecture, sWindowsVersion ' Windows version and architecture - set in WinArch() and WinVer()
Dim sUserPath, sProgramFiles, sSoftwareRegistryHive ' Windows Variables based on version and architecture - set in WindowsCheck()

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