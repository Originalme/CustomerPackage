TestExist = CheckFile("..\..\bin\Test.cmd")

Function CheckFile(sFile)

	dim oFileSys
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	
	'oFileSys.CreateTextFile sFile, True
	
	If oFileSys.FileExists(sFile) Then
		CheckFile = True
	Else
		CheckFile = False
	End If 

End Function

MsgBox TestExist