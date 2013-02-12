'------------------------------------------------------'
'  Function: Get Windows Version baed on Build Number  '
'------------------------------------------------------'
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