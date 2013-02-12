'----------------------------------------------------------------------------'
'                          Check MSI version                                 '
'  You can pass an MSI file to this function to get the version of the MSI.  '
'----------------------------------------------------------------------------'
Function GetMsiVersion(sMSIFile)
	Dim  Installer, Database, SQL, View, Record
	
	Set Installer = CreateObject("WindowsInstaller.Installer")
	
	Set Database = Installer.OpenDatabase(sMSIFile, 0)
	
	SQL = "SELECT * FROM Property WHERE Property = 'ProductVersion'"
	
	Set View = DataBase.OpenView(SQL)
	
	View.Execute
	
	Set Record = View.Fetch
	
	GetMsiVersion = Record.StringData(2)
	'MsgBox "MSI File Version = " & sMSIFile

End Function