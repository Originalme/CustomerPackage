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