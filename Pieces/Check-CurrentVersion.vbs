Dim sCheckRegHive, sCitrixRecNewVersion, sViewNewVersion, sAMNewVersion, sOneSignNewVersion, sIMRDPNewVersion ' Versions of MSI's contained in package and IM-OneInstaller Reg Hive - set in UpdateCheck()

'----------------------------------------------------------------------------------------------------------------------------------------------------'
'                                         !!!Primary Function - Check For Updates!!!                                                                 '
'----------------------------------------------------------------------------------------------------------------------------------------------------'
'  Primary Function: Check for First Run. Then finds and compares the versions of the packeges currently installed vs. the ones in the new package.  '
'----------------------------------------------------------------------------------------------------------------------------------------------------'
'    If new packages are found update proceeds.  '
'    If no packages found, installer runs        '
'    If no updates found, installer exits        '
'------------------------------------------------'
Function UpdateCheck() ' Checks to see if this is the first time installing the package, or if there are any updates to be made.
	Dim sFirstRun
	Dim sCitrixRecCurrentVersion, sViewCurrentVersion, sAMCurrentVersion, sOneSignCurrentVersion, sIMRDPCurrentVersion
	
	'-----------------------------------------------------------------'
	'                             First Run                           '
	'  Checks for the IM-OneInstaller Reg Key, if not exist make it   '
	'-----------------------------------------------------------------'
	WriteLog(Chr(13) & "--------------------------------Checking First Run--------------------------------" & Chr(13))
	sCheckRegHive = "HKEY_LOCAL_MACHINE\SOFTWARE\IM-OneInstaller\"
	
	sFirstRun = KeyExists(HKLM, "SOFTWARE\IM-OneInstaller\")
	
	if sFirstRun = "False" Then
		
		WriteLog("Running as first time installation")
		WriteLog("Creating registry hive... " & sCheckRegHive)
		
		CreateRegKey(sCheckRegHive)
	
	Else

	WriteLog("Registry hive already found... " & sCheckRegHive)
		WriteLog("Running as an update")
	
	End if
	
	'-------------------------------------------------------------------------'
	'                    Check MSI File Versions                              '
	'  Checks OneSign and Authentication Manager MSI versions in the package  '
	'-------------------------------------------------------------------------'
	WriteLog(Chr(13) & "----------------------------Checking Packaged Versions-----------------------------------" & Chr(13))
	If "sWindowsArchitecture" = "x64" Then
		sOneSignNewVersion = GetMsiVersion(sBinDir & "OneSignAgentx64.msi")
		WriteLog("OneSign x64 MSI version... " & sOneSignNewVersion)
	
	Else 
		sOneSignNewVersion = GetMsiVersion(sBinDir & "OneSignAgent.msi")
		WriteLog("OneSign x86 MSI version... " & sOneSignNewVersion)
	End If
	
	sAMNewVersion = GetMsiVersion(sBinDir & "AMCLIENT.msi")
	WriteLog("Authentication Manager MSI version... " & sAMNewVersion)
	
	'-------------------------------------------------------------------------'
	'                    Check EXE File Versions                              '
	'  Checks XenApp, XenDesktop, and VMwareView EXE versions in the package  '
	'-------------------------------------------------------------------------'
	If "sWindowsArchitecture" = "x64" Then
		sViewNewVersion = GetExeVersion(sBinDir & "VMWare\ViewClientx64.exe")
		WriteLog("View Client x64 EXE version = " & sViewNewVersion)
	Else
		sViewNewVersion = GetExeVersion(sBinDir & "VMWare\ViewClient.exe")
		WriteLog("View Client x86 EXE version = " & sViewNewVersion)
	End If
	
	sCitrixRecNewVersion = GetExeVersion(sBinDir & "CitrixReceiver.exe")
	WriteLog("Citrix Receiver EXE version... " & sCitrixRecNewVersion)

	sIMRDPNewVersion = GetExeVersion(sBinDir & "IMRDP\IMONERDP.exe")
	WriteLog( "IM-RDP EXE Version = " & sIMRDPNewVersion)
	
	'----------------------------------------------------------'
	'               Check Current Installed Version            '
	'  Checks the currently installed version of all software  '
	'----------------------------------------------------------'
	WriteLog(Chr(13) & "------------------------------Checking Currently Installed Versions---------------------------------------" & Chr(13))
	
	sOneSignCurrentVersion = ReadRegValue(sCheckRegHive & "OneSignVer", 0)
	WriteLog("OneSign installed version... " & sOneSignCurrentVersion)
	
	sCitrixRecCurrentVersion = ReadRegValue(sCheckRegHive & "CTXVer", 0)
	WriteLog("Citrix Man installed Version... " & sCitrixRecCurrentVersion)
	
	sViewCurrentVersion	= ReadRegValue(sCheckRegHive & "VMWareVer", 0)
	WriteLog("VMWare installed Version... " & sViewCurrentVersion)
	
	sAMCurrentVersion = ReadRegValue(sCheckRegHive & "AuthManVer", 0)
	WriteLog("Auth Man installed Version... " & sAMCurrentVersion)
	
	sIMRDPCurrentVersion = ReadRegValue(sCheckRegHive & "IM-RDPVer", 0)
	WriteLog("IMRDP installed Version... " & sIMRDPCurrentVersion)
	
	' Check for Updates based on registry keys if none are found exit.
	If sOneSignNewVersion <> sOneSignCurrentVersion _
	OR sCitrixRecNewVersion <> sCitrixRecCurrentVersion _
	OR sViewCurrentVersion <> sViewNewVersion _
	OR sAMCurrentVersion <> sAMNewVersion _
	OR sIMRDPCurrentVersion <> sIMRDPNewVersion _
	Then
		WriteLog ("Update Found, proceeding with installation")
	Else	

		WriteLog("No updates found")
		WriteLog(" Exiting installer")
		
		MsgBox "There are no updates at this time."
		Wscript.Quit
	End If
	
End Function