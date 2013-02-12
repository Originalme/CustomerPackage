'------------------------------------------------------------------'
'                     Get Registry Key Value                       '
'  Passing a registry key to this function will return its value.  '
'------------------------------------------------------------------'
Function ReadRegValue(sRegKey, sDefault)
	Dim wshShell, value

	On Error Resume Next
	Set wshShell = CreateObject("WScript.Shell")
	value = wshShell.RegRead(sRegKey)

	if err.number <> 0 then
		WriteLog("Key not found... " & sRegKey)
		ReadRegValue= sDefault
	else
		WriteLog("Found Value... " & sRegKey)
		ReadRegValue = wshShell.RegRead(sRegKey)
	end if
	
	set wshShell = nothing
	
End Function