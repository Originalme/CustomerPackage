'-----------------------------------------'
'  Function: Get Windows OS Architecture  '
'-----------------------------------------'
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