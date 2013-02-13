Option Explicit

Const sBinDir = "C:\Users\chrisb\GitHub\CustomerPackage\CustomerPackage\Bin\"
Const sComputer = "."
Const ALL_USERS = True





InstallEXE "AMCLIENT.msi", ""

Sub InstallEXE(sExecutable, sCommandLine1)
	Dim oSoftware, oService, errReturn
	
	sExecutable = Chr(34) & sBinDir & sExecutable & Chr(34)
	MsgBox sExecutable
	Set oService = GetObject("winmgmts:"  & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
	Set oSoftware = oService.Get("Win32_Product")
	
	errReturn = oSoftware.Install(sExecutable,, ALL_USERS)
	
	MsgBox errReturn
	
	
End Sub