Windows Support Tool Kit
========================

[Capac Techne IT Services](https://www.facebook.com/pg/Capac-Techne-Servicios-Informatico-253256362049018) brings you SupportTK, the Support Tool Kit Powershell v5 module, to assist technical support engineers in various efforts, from data gathering, to diagnosis and remediation steps. 

Purpose
-------

Despite significant improvements in Windows 10 and Windows Server 2016, there are 
still gaps in command-line tools for 1st tier support, or those who are not well 
versed in PowerShell. Also, level 1 support teams lack the mandate and resources 
to develop a powershell library of advanced functions to cover basic use cases. 

The Support toolkit provides functions to collect information quickly, so support 
staff can devote more time to trouble shooting and diagnosis. 

SupportTK provides task-based commands for basics like:

- list patches installed (get-patch)
- determine environment and version details for Windows OS (get-windows)
- scan Windows Event Log for key entries based on known error IDs (coming soon)
- determine degree of disk fragmentation
- folder size and growth
- determine locked files & unlock these by closing app (remove-lockedfile)
- locate multiple copies of executables in the path
- isolate DNS issues related to DC and GC servers (coming soon)



Advanced Functions
------------------

- Get-EnvPath: Displays directories, one per line, declared in user's path environment variable.
- Get-Patch: Lists hotfixes installed on computer.
- Get-LockedFile: Displays applications with opened or locked files in local file system.
- Get-TempDir: Returns file object for Temp directory, based on type, and basic statistics.
- Get-Windows: Identifies Windows Operating System: Maj.Min.Build.Release and Update Numbers
- New-TempDir: Creates sub directory with random name within temp folder on SYSTEMDRIVE.
- Remove-LockedFile: Closes application holding locked file, then deletes the file.
- Enable-LockedFile: Controls the 'maintain objects list' so openfile can track handles.
- Test-EnvPath: Determines directory on the path wherein the specified file resides.


Package Requirement
-------------------

Windows PowerShell 4.0 or higher
openfiles.exe (Windows 8 or higher)


License
-------

Released under GPL 2.0





