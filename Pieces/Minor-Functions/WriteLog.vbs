'---------------------------'
'         Write Log         '
'  Writes line to logfile.  '
'---------------------------'
Function WriteLog(sLogLine)
	Dim oFileSystem, oFile
	set oFileSystem = CreateObject("Scripting.FileSystemObject")
	set oFile = oFileSystem.OpenTextFile(sInstallLog, ForAppending, True)

	oFile.WriteLine(Now & "  |  " & sLogLine)

	oFile.Close

End Function