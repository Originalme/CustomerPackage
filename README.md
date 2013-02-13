README for IM-One Installer/Customer Package
Written By: Christopher S. Bates

Rules for Editing
------------------

1. DO NOT directly edit the IM-OneInstaller.vbs.
	* Any comits with changes to this file will be rejected immediately.
	* Please modify the function files and commit them with an explanation of where they should go in the main file.

2. Modification/Removal of any function or sub function must be justified and explained in the git commit comment
	* If you are going to work on a new Primary function, please create a blank file to claim it and performa  pull request

3. Don't annoy the repo owner.

Bin Directory
-------------

* The binary files are not included in this repo. To obtain these you must be a current employee of FAI and request access
	to them. You may ask Chris Bates, as he SHOULD know where they are stored.

Workflow Diagram
-----------------

* Use this to guide you as you are working. If you wish to add something please modify it as a seperate branch, and commit it alone.
* Any changes will be reviewed by the team before being merged.

---------------REPO TREE--------------------------
Folder PATH listing
Volume serial number is BAFA-7CA7
C:.
|   .gitattributes
|   .gitignore
|   ChangeLog.txt
|   IM-OneInstaller.vbs
|   InstallLog.txt
|   KnownIssues.txt
|   README.md
|   README_2.md
|   WorkflowDiagram.vsdx
|   
+---Bin
|   |   AMCLIENT.msi
|   |   CitrixOnlinePluginFull.exe
|   |   CitrixReceiver.exe
|   |   OneSignAgent.msi
|   |   OneSignAgentx64.msi
|   |   
|   +---IMRDP
|   |       ChangeLog.TXT
|   |       IMONERDP.EXE
|   |       IMONERDP.INI
|   |       IMONERDPCB.EXE
|   |       IMONERDPCBDBG.EXE
|   |       IMONERDPDBG.EXE
|   |       IMONERDP_DEBUG.log
|   |       launchcbdbg.cmd
|   |       
|   +---VMWare
|   |       ViewClient.exe
|   |       ViewClientx64.exe
|   |       
|   \---XenApp
|       \---FastConnect
|           |   Configuring XenApp Roaming Using the Citrix Fast Connect Utility.pdf
|           |   FastConnect-Disconnect.txt
|           |   FastConnect-Launch.txt
|           |   Kill Process (WScript) on Lock.txt
|           |   Lock on WFICA32(32) Exit.txt
|           |   
|           +---x64
|           |       citrixfastconnect.exe
|           |       fastconnect.dll
|           |       
|           \---x86
|                   citrixfastconnect.exe
|                   fastconnect.dll
|                   
\---Pieces
    |   Check-CurrentVersion.vbs
    |   DisableUAC.vbs
    |   GlobalConstants.vbs
    |   WindowsCheck.vbs
    |   
    +---Minor-Functions
    |       CheckFileVersion.vbs
    |       CreateRegistryKey.vbs
    |       DeleteFile.vbs
    |       ErrorDialogueBox.vbs
    |       GetMSIVersion.vbs
    |       GetWindowsArchitecture.vbs
    |       GetWindowsVersion.vbs
    |       ReadRegValue.vbs
    |       RegKeyExists.vbs
    |       SetDWORDRegistry.vbs
    |       SetStringRegistry.vbs
    |       WriteLog.vbs
    |       
    \---Notes
            Check-CurrentVersion-Notes.txt
            WindowsCheck-Notes.txt
            
