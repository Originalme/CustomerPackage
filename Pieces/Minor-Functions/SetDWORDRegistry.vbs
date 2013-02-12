'----------------------------------------------'
'                 Set DWORD Key                '
'  Sets a DWORD Value in the Windows Registry  '
'----------------------------------------------'
Function SetRegDword(sKeyHive, sKeyPath, sValueName, sValue)
	Dim oRegistry
	Set oRegistry = GetObject("winmgmts:\\" & sComputer & "\root\default:StdRegProv")
	
	oRegistry.SetDWORDValue sKeyHive, sKeyPath, sValueName, sValue
	

End Function
