Option Explicit
'------------------------------------------------------------------------'
'                           Disable UAC                                  '
'  If Windows 7 or Windows Vista is detected, this will disable the UAC  '
'  Written By: Christopher S. Bates                                      '
'------------------------------------------------------------------------'
'  Version 0.5.0  '
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
Const sInstallLog = "..\InstallLog.txt"								'Location of the install log
Const ForAppending = 8


Dim sWindowsVersion: sWindowsVersion = 7601
'--------------------------------------------------------------------END OF GLOBAL--------------------------------------------------------------------'
'-------Dependencies-------'
'	  WindowsCheck.vbs     '
'-------Dependencies-------'

If sWindowsVersion = 6000 _
OR sWindowsVersion = 6002 _
OR sWindowsVersion = 7600 _
OR sWindowsVersion = 7601 _
Then
	Dim sResponse
	
	sResponse = MsgBox("Would you like to disable the UAC?", 4, sAppName) 'User selects to disable UAC or not
	
	If sResponse = 6 Then
		DisableUac()
	Else
		WriteLog("User Action: Do not disable UAC")
	End If
	
End If
	
	
	
Sub DisableUac()
	Dim sUACRegValue, CommandLine
	
	
	sUACRegValue = ReadRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA", 1)
	
	If  sUACRegValue <> 0 Then 
		WriteLog("Disabling UAC")
		SetRegDword HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA", 0
		
	End If
	
	sUACRegValue = ReadRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA", 1) 'Logging for UAc
	WriteLog("UAC Value set to... " & sUACRegValue)
	
End Sub
	
	
	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------SUB FUNCTIONS-----------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'

'------------------------------------------------------------------'
'                        Get Registry Key Value                    '
'  Passing a registry key to this function will return its value.  '
'------------------------------------------------------------------'
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

'----------------------------------------------------------'
'                     Create Registry Key                  '
'  Imput full registry key path and this will generate it  '
'----------------------------------------------------------'

Function CreateRegKey(sKey)
	Dim wshShell
	Set wshShell = CreateObject( "WScript.Shell" )
	
	wshShell.RegWrite sKey, ""
	
	Set wshShell = nothing

End Function

'----------------------------------------------'
'                 Set DWORD Key                '
'  Sets a DWORD Value in the Windows Registry  '
'----------------------------------------------'
Function SetRegDword(sKeyHive, sKeyPath, sValueName, sValue)
	Dim oRegistry
	Set oRegistry = GetObject("winmgmts:\\" & sComputer & "\root\default:StdRegProv")
	
	oRegistry.SetDWORDValue sKeyHive, sKeyPath, sValueName, sValue
	

End Function


'-----------------------------------------------'
'                Set String Key                 '
'  Sets a String Value in the Windows Registry  '
'-----------------------------------------------'
Function SetRegString(sKeyHive, sKeyPath, sValueName, sValue)
	Dim oRegistry, oKeyType
	Set oRegistry = GetObject("winmgmts:\\" & sComputer & "\root\default:StdRegProv")
	
	
	oRegistry.SetStringValue sKeyHive, sKeyPath, sValueName, sValue
	

End Function

'---------------------------'
'         Write Log         '
'  Writes line to logfile.  '
'---------------------------'
Function WriteLog(sLogLine)
	Dim oFileSystem, oFile
	set oFileSystem = CreateObject("Scripting.FileSystemObject")
	set oFile = oFileSystem.OpenTextFile(sInstallLog, ForAppending, True)

	oFile.WriteLine(Now & "  |  " & sLogLine)

	oFile.Close

End Function
	
	
	