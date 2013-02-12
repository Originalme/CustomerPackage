'-----------------------------------------------'
'                Set String Key                 '
'  Sets a String Value in the Windows Registry  '
'-----------------------------------------------'
Function SetRegString(sKeyHive, sKeyPath, sValueName, sValue)
	Dim oRegistry, oKeyType
	Set oRegistry = GetObject("winmgmts:\\" & sComputer & "\root\default:StdRegProv")
	
	
	oRegistry.SetStringValue sKeyHive, sKeyPath, sValueName, sValue
	

End Function
