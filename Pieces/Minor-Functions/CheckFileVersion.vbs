'----------------------------------------------------------------------------'
'                            Check EXE version                               '
'  You can pass an EXE file to this function to get the version of the EXE.  '
'----------------------------------------------------------------------------'
Function GetExeVersion(sExeVersion)
	
	'MsgBox sExeVersion
	GetExeVersion = CreateObject("Scripting.FileSystemObject").GetFileVersion(sExeVersion)

End Function