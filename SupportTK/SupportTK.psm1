<# 
    .SYNOPSIS   
    Tools to enhance technical assistance efforts for Windows Desktops.

    .DESCRIPTION 
    Collection of task-based advfuncs oriented to tech support activities and diagnostics.

    .NOTES   
    Requires openfiles.exe on VM or remote workstation
    Owner       : Capac Techne
    Module      : SupportTK
    Designer    : Sean Peterson
    Contributors: 
    Created     : 2018-08-22
    Updated     : 2018-10-13
    Version     : 0.10.7
    2018-10-17  : Added function to Get-Patch to list hotfixes installed in the last 45 days

#>

Function Get-Patch { # Displays installed patches (hotfixes).
	[CmdletBinding(DefaultParametersetName="Standard")]
	Param(
      [Parameter(ParameterSetName='Standard',Position=0)]
      [Alias("C","Days")]
      [STRING]$CutOff = 45,
      
      [Parameter(ParameterSetName='Standard',Position=1)]
      [ValidateSet('InstalledOn','Name')]
      [Alias("S","SortOn")]
      [STRING]$SortBy = "Installed"
    ) 
 
    if ($SortBy -eq "Name") {$SortBy = "HostfixID"}
    Get-Hotfix | Where {$_.InstalledOn -gt $(Get-Date).AddDays(-${cutoff})} | Sort-Object -Property $SortBy
}
Set-Alias showpatch Get-Patch -Description "Displays installed patches (hotfixes)."  


Function Get-EnvPath { # Displays directories, one per line, declared in user's path environment variable.
    $env:Path.Split(';')
}
Set-Alias showpath Get-EnvPath -Description "Displays the directories, one per line, of the path."  


Function Test-EnvPath { # Determines directory on the path wherein the specified file resides. 
	[CmdletBinding(DefaultParametersetName="Standard")]
	Param(
	  [Parameter(ParameterSetName='Standard',Position=0)][STRING]$File,
	  [Parameter(ParameterSetName='Standard',Position=1)][SWITCH]$Detail
    ) 
    
    $pathsplit = $env:Path.Split(';')
    If ($File) {
        # Search for the specified file in each directory on path
        ForEach ($directory in $PathSplit) {
            Get-ChildItem -Path $directory -Filter $File -EA 0; 

        }
    } else {
        # Just list each directory in the path, indicating whether it exists in file system
        ForEach ($directory in $PathSplit) {

            $Result = Test-Path $directory 
            $Props = @{ Path = $directory; Exists = $Result }
            New-Object -Type PSObject -Property $Props

        }
    }
}
Set-Alias which Test-EnvPath -Description "Determines where in the path the specified file resides."  


Function New-TempDir { # Creates sub directory with random name within temp folder on SYSTEMDRIVE. 

    $parent     = [System.IO.Path]::GetTempPath()
    $filename   = [System.IO.Path]::GetRandomFileName()
    $newtempdir = Join-Path $parent $filename
    New-Item -ItemType Directory -Path $newtempdir
    
}
Set-Alias mktmpdir New-TempDir -Description "Creates sub directory with random name in temp folder."  


Function Get-TempDir { # Returns file object for Temp directory, based on type, and basic statistics.
    [CmdletBinding(DefaultParametersetName="Standard")]
    Param(
        [Parameter(
            ParameterSetName='Standard',
            Position=0,
            Mandatory=$False
        )][ValidateNotNullOrEmpty()]
        [Alias("Dir","Path")]
        [STRING]$Folder=$($env:TMP),

        [Parameter(
            ParameterSetName='Advanced',
            Position=0,
            Mandatory=$False
        )][ValidateNotNullOrEmpty()]
        [Alias("SB")]
        [STRING]$Type,
        
        [Parameter(ParameterSetName='Standard',Position=1)]
        [Parameter(ParameterSetName='Advanced',Position=1)]
        [SWITCH]$Stats
    )
  
    Begin{

        If ($Type) {
            Switch ($Type) {

                "System" { # gci on the current working directory
                    $TargetDir = Get-Item -Path $([System.IO.Path]::GetTempPath())
                }

                "CWD" { # gci on the current working directory
                    $TargetDir = Get-Item -Path ".\"
                }

                "Windows" { # gci on the current working directory
                    $TargetDir = Get-Item -Path "$($env:windir)\Temp"
                }

                Default { # gci on a profile folder based on UserProfile environment var
                    $TargetDir = Get-Item -Path (Join-Path (Get-Item -Path ($env:UserProfile)).Parent.Fullname "${Type}\AppData\Local\Temp")
                }
            }
        }
        ElseIf ($Folder -and (Test-Path $Folder)) {
            $TargetDir =  get-Item -Path $Folder
        }
        Else {
            Throw "Folder $Folder not found."
        }

        $PathTmp = $TargetDir;
    }
    Process{

        $PathTmp
        if ($Stats) {
            ${StatsInfo} = Get-ChildItem -Path $pathTmp -Recurse ;
            ${StatsInfo} | Measure-Object -Property length -Minimum -Maximum -Average -Sum |
                Select-Object -Property @{name="File Count";e={$_.Count}},
                    @{name="Smallest (MB)";e={[math]::round(($_.Minimum/1MB),2)}},
                    @{name="Largest (MB)";e={[math]::round(($_.Maximum/1MB),2)}},
                    @{name="Average (MB)";e={[math]::round(($_.Average/1MB),2)}},
                    @{name="Total (MB)";e={[math]::round(($_.Sum/1MB),2)}}
        }
        
    }
    End {}

}
Set-Alias tmpdir Get-TempDir -Description "Displays location of temp directory."


Function Stop-AppLocking { # Place-holder Function
    [CmdletBinding(DefaultParametersetName="Standard")]
    Param(
        [Parameter(
            ParameterSetName='Standard',
            Position=0,
            Mandatory=$True,
            ValueFromPipeline=$True
        )]
        [STRING]$ID,
        [Parameter(
            ParameterSetName='Standard',
            Position=1,
            Mandatory=$True,
            ValueFromPipeline=$True
        )]
        [STRING]$PID,
        [Parameter(
            ParameterSetName='Standard',
            Position=2,
            Mandatory=$True,
            ValueFromPipeline=$True
        )]
        [STRING]$FilePath        
    )
    
    # uses 
    
    stop-process -ID $PID -force 

}


Function Remove-LockedFile { # Safely removes lock and deletes a file ("delflock"). 

    [CmdletBinding(DefaultParametersetName="Standard")]
    Param(
        [Parameter(
            ParameterSetName='Standard',
            Position=0,
            Mandatory=$True,
            ValueFromPipeline=$True
        )]
        [STRING]$FilePath,

        [Parameter(ParameterSetName='Standard',Position=1,Mandatory=$True)]
        [STRING]$Cached,

        [Parameter(
            ParameterSetName='Release',
            Position=0,
            Mandatory=$True,
            ValueFromPipeline=$True            
        )][Alias("HandleId")]
        [INT]$ID
    )

    Begin{

        $InformationPreference = "SilentlyContinue"; 

        If ( ${cached} -and (-not (Test-Path ${cached})) )  {
            Remove-Variable cached
            "Cache is missing at $Cached"
            Break
        }
    }
    Process{
        If ( ${ID} ) {
            $R = Get-LockedFile -Id ${ID}
        } 
        ElseIf ( ${filePath} -and ${cached} ) {
            $R = Get-LockedFile -filter "${filePath}" -cached "${cached}" #| Format-Table -Property ID,FilePath,ProcessName -auto -wrap
        } 
        Else {
            $R = Get-LockedFile -filter "${FilePath}"
        }

        $S = @($R.PID)
        
        If ($S.Count -eq 1) {
            
            $LockingApp = Get-Process -Id $R.PID
            $LockingApp | Select Name, Id, Product, ProductVersion, Path

            $prompt = Read-Host -Prompt "Terminate process $($LockingApp.ID)`? yes / No";
            If ($prompt -eq "Yes") {
                
                # Stop-AppLocking -id $($R.ID) -Pid $($R.PID) -File "$($R.FilePath)"
                Write-Information "Waiting to unlock file $($R.FilePath)`. "
                Stop-Process -Id $R.PId
                Start-Sleep -s 2
                Remove-Item $R.FilePath
                if ( Test-Path $R.FilePath ){ 
                    Write-Information "File not deleted, still locked. Try again.";
                } 
                Else {
                    Write-Information "Deleted locked file $($R.FilePath)`. ";
                }
                
            } 
            Else {
                Write-Information "Selected option 'No' (default)."
            }
            
        } 
        Else {
            Write-Error "No single valid file handle found. Handle ${ID} appears $($S.Count) times."
        }
    }
    End {}

}
Set-Alias delflock Remove-TempFile -Description "Displays X."

#deltmp -filepath X

Function Enable-LockedFile { # Controls the 'maintain objects list' so openfile can track handles.
    [CmdletBinding(DefaultParametersetName="Active")]
    Param(
        [Parameter(ParameterSetName='Active',Position=0)][SWITCH]$Enabled,
        [Parameter(ParameterSetName='Inactive',Position=0)][SWITCH]$Disabled
    )
  
    Begin{ }

    Process{
        if ($Enabled) {
            $message = openfiles /local ON
        } elseif ($Disabled) {
            $message = openfiles /local OFF
        } else {
            $message = openfiles /local
        }
    }

    End{
        Write-Verbose "$message"
    }

}
Set-Alias trackflock Set-OpenFile -Description "Controls the 'maintain objects list' so openfile can track."


Function Get-LockedFile { # Displays applications with opened or locked files in local file system.
    [CmdletBinding(DefaultParametersetName="Standard")]
    Param(
        [Parameter(
            ParameterSetName='Standard',
            Position=0,
            ValueFromPipeline=$True
        )]
        [Alias("FilePattern")]
        [STRING]$Filter,

        [Parameter(ParameterSetName='Standard',Position=1)]
        [STRING]$Cached,

        [Parameter(ParameterSetName='Standard',Position=0)]
        [Alias("HandleId")]
        [INT]$ID
    )
  
    Begin{

        $InformationPreference = "SlientlyContinue";
        ${FileCache}= ${Cached};
        ${StartStr} = "Files Opened Locally:";
        ${EndStr}   = "Files opened remotely via local share points:";
        ${Header}   = "ID","AccessedBy","PID","ProcessName","FilePath";
        ${Pattern}  = $Filter; 

    }
    Process{

        Try {

            If ( ${FileCache} -and (Test-Path ${FileCache}) ) {                

                ${FileObj} = Get-Item ${FileCache};
                ${Message} = "Using cached data in ${FileCache} generated $(${FileObj}.LastWriteTime)";
                ${TextRaw} = Get-Content ${FileCache};

            } Else {

                Write-Information "No cache found at ${FileCache}, so scanning operation may take several minutes to complete. `nPlease do not close this session. "
                $TimeResult = Measure-Command -Expression { 
                    ${TextRaw} = openfiles /query /v /fo CSV /nh;
                } 
                If (${FileCache}) { ${TextRaw} > ${FileCache} }
                ${Message} = "Openfiles completed in $($TimeResult.TotalMinutes) minutes."

            }

            $OpenFiles = ${TextRaw}[($TextRaw.IndexOf(${StartStr}) + 2) .. ($TextRaw.IndexOf(${EndStr}) -3)] | ConvertFrom-Csv -Header ${Header};

        } 
        Catch {
            Write-Output $Error[0]
        }
        Finally{
            Write-Verbose ${Message}
        }

        If ( $ID ) {
            $OpenFiles | Where-Object { $_.ID -eq "${ID}" } 
        } ElseIf ( $Filter ) {
            #$FinalResult = 
            $OpenFiles | Where-Object { $_.FilePath -Like "*${Pattern}*" }
        } Else {
            #$FinalResult = ${OpenFiles}
            $OpenFiles 
        }

    }
    End{
        #$FinalResult | Format-Table -Property ID,FilePath,PID,ProcessName -auto -wrap
    }

}
Set-Alias flock Get-LockedFile -Description "Displays applications with opened or locked files in local file system."


Function Get-Windows { # Identifies Windows Operating System: Maj.Min.Build.Release.Update

    [CmdletBinding(DefaultParameterSetName="Version")]
    Param(
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='MajorNo')][SWITCH]$Major,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='MinorNo')][SWITCH]$Minor,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='BuildNo')][SWITCH]$Build,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='RevisionNo')][SWITCH]$Revision,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='ReleaseID')][SWITCH]$Release,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='Role')][SWITCH]$Role,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='Product')][SWITCH]$Product,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='ProductVer')][SWITCH]$ProductVersion,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='VersionObj')][SWITCH]$Version,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='Serial')][SWITCH]$SerialNumber,
        [Parameter(Position=0,Mandatory=$False,ParameterSetName='Up')][SWITCH]$Uptime
    )

    Begin {
        
        $w32OS = Get-CimInstance -ClassName win32_operatingsystem -EA 0;
        $UBR = [STRING](Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion" UBR).UBR
        $ReleaseId = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId).ReleaseId;

        $Ver = $w32OS.Version.ToString().Split(".");
        $Ver += $UBR;

        $objVersion = New-Object -typename System.Version $Ver;
        $strVersion = $Ver -Join "."

        $DomainRole = @{
            0 = "Standalone Workstation" ;
            1 = "Member Workstation" ;
            2 = "Standalone Server" ;
            3 = "Member Server" ;
            4 = "Backup Domain Controller" ;
            5 = "Primary Domain Controller";
            6 = "Read-Only Domain Controller" 
        } 

    }
    Process{
        
        switch($PSCmdlet.ParameterSetName){
            
            "MajorNo"   { $objVersion.Major; Break; }

            "MinorNo"   { $objVersion.Minor; Break; }

            "BuildNo"   { $objVersion.Build; Break; }

            "Product"   { $w32OS.Caption; Break; }

            "ProductVer"{ "$($w32OS.Caption), Release $ReleaseID [Build $($Ver[2])`.$UBR]"; Break; }

            "RevisionNo"{ $objVersion.Revision; Break; }

            "ReleaseID" { $ReleaseId; Break; }

            "Role"      { $w32CS = Get-CimInstance -ClassName win32_computersystem -EA 0; 
                            $DomainRole[[int]$w32CS.DomainRole]; 
                            Break; 
                        }

            "Serial"    { $w32OS.SerialNumber; Break; }

            "Up"        { $when = $w32OS.LastBootUpTime.GetDateTimeFormats()[112];
                            $T = (Get-Date).Subtract($w32OS.LastBootUpTime);
                            $U = $T.Subtract($T.Days); 
                            "$($w32OS.CSName.ToLower()) $When up $($T.Days) days, $($U.Hours)`:$($U.Minutes)"; 
                            Break; 
                        }

            "VersionObj" { $objVersion; Break; }
            
            Default      { $strVersion }

        }
    }
    End{}

    
    <#

        .SYNOPSIS
        Displays the Windows Operating System number format: Major.Minor.Build.Revision

        .DESCRIPTION
        Displays version of Windows Operating System: Major Version, Minor Version, 
        Build Number, Revision Number. Can also return the semi-annual Release ID. 
        N.B. The revision number reflects the UBR (Updated Build Release Identification Number).

        .INPUTS
        This cmdlet accepts no pipeline input

        .OUTPUTS
        The four-part version number of Windows OS.

        .EXAMPLE
        Get-Windows -ProductVersion
        Microsoft Windows 10 Home, Release 1803 [Build 17134.285]

        .EXAMPLE
        Get-Windows -Version
        10.0.17134.285

        .EXAMPLE
        Get-Windows -Release
        1803

        .EXAMPLE
        Get-Windows -Role
        Standalone Workstation

        .EXAMPLE
        Get-Windows -Product
        Microsoft Windows 10 Home

        .EXAMPLE
        Get-Windows -SerialNumber
        000000-00000-00000-AAOEM

        .EXAMPLE
        Get-Windows -Uptime
        15:31:19 up 0 days, 10:12

    #>

}
Set-Alias version Get-Windows -Description "Identifies Windows Operating System: Maj.Min.Build.Release.Update."



$smasktoCIDR = @{
        '255.0.0.0'       = [byte]'8' 
        '255.128.0.0'     = [byte]'9' 
        '255.192.0.0'     = [byte]'10'
        '255.224.0.0'     = [byte]'11'
        '255.240.0.0'     = [byte]'12'
        '255.248.0.0'     = [byte]'13'
        '255.252.0.0'     = [byte]'14'
        '255.254.0.0'     = [byte]'15'
        '255.255.0.0'     = [byte]'16'
        '255.255.128.0'   = [byte]'17'
        '255.255.192.0'   = [byte]'18'
        '255.255.224.0'   = [byte]'19'
        '255.255.240.0'   = [byte]'20'
        '255.255.248.0'   = [byte]'21'
        '255.255.252.0'   = [byte]'22'
        '255.255.254.0'   = [byte]'23'
        '255.255.255.0'   = [byte]'24'
        '255.255.255.128' = [byte]'25'
        '255.255.255.192' = [byte]'26'
        '255.255.255.224' = [byte]'27'
        '255.255.255.240' = [byte]'28'
        '255.255.255.248' = [byte]'29'
        '255.255.255.252' = [byte]'30'
    }