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