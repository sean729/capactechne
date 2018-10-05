Windows Support Tool Kit
========================

[Capac Techne IT Services](https://www.facebook.com/pg/Capac-Techne-Servicios-Informatico-253256362049018) brings you SupportTK, the Support Tool Kit Powershell v5 module, to assist technical support engineers in various efforts, from data gathering, to diagnosis and remediation steps. 



Advanced Functions
------------------

- Get-EnvPath: Displays directories, one per line, declared in user's path environment variable.
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


Purpose
========

Despite significant improvements in Windows 10 and Windows Server 2016, there still exist gaps in the command-line offering in terms of what information is readily available when support staff perform basic information collection and diagnosis. Generally, level 1 support doesn't have the mandate or resources to invest in developing a powershell tool kit that covers such use cases. 

SupportTK provides support staff the advanced functions to perform certain basics, like:

- log scrubbing
- disk fragmentation
- folder size and growth
- determine locked files (& unlock these)
- locate multiple copies of exe in the path
- isolate DNS issues related to DC and GC servers




