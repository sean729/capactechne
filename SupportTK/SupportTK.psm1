<# 
    .SYNOPSIS   
    Tools to enhance technical assistance efforts on Windows Desktops.

    .DESCRIPTION 
    Collection of task-based advfuncs oriented to tech support activities and diagnostics.

    .NOTES   
    Requires openfiles.exe on VM or remote workstation
    Owner       : Capac Techne
    Module      : SupportTK
    Designer    : Sean Peterson
    Contributors: 
    Created     : 2018-08-22
    Updated     : 2018-09-16
    Version     : 0.8
    2018-09-16  : tech support adv functions Get-EnvPath, Test-EnvPath, New-TempDir, Get-TempDir

#>

Function Get-EnvPath { # Displays directories, one per line, declared in user's path environment variable.
    $env:Path.Split(';')
}; Set-Alias showpath Get-EnvPath -Description "Displays the directories, one per line, of the path."  


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

}; Set-Alias which Test-EnvPath -Description "Determines where in the path the specified file resides."  


Function New-TempDir { # Creates sub directory with random name within temp folder on SYSTEMDRIVE. 

    $parent     = [System.IO.Path]::GetTempPath()
    $filename   = [System.IO.Path]::GetRandomFileName()
    $newtempdir = Join-Path $parent $filename
    New-Item -ItemType Directory -Path $newtempdir
    
}; Set-Alias mktmpdir New-TempDir -Description "Creates sub directory with random name in temp folder."  


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

}; Set-Alias tmpdir Get-TempDir -Description "Displays location of temp directory."


Function Remove-OpenFile { # Place-holder Function
    [CmdletBinding(DefaultParametersetName="Standard")]
    Param(
        [Parameter(
            ParameterSetName='Standard',
            Position=0,
            Mandatory=$True,
            ValueFromPipeline=$True
        )]
        [STRING]$ID
    )
    
    # uses stop-process to unlock local bin files

}


Function Remove-LockedFile { # Safely removes lock and deletes a file ("delflock"). 
    # BTR 

    [CmdletBinding(DefaultParametersetName="Standard")]
    Param(
        [Parameter(
            ParameterSetName='Standard',
            Position=0,
            Mandatory=$True,
            ValueFromPipeline=$True
        )]
        [STRING]$FilePath,

        [Parameter(ParameterSetName='Standard',Position=1)]
        [STRING]$Cached,

        [Parameter(
            ParameterSetName='Release',
            Position=0,
            Mandatory=$True,
            ValueFromPipeline=$True            
        )]
        [INT]$ID
    )

    If ( ${cached} -and (-not (Test-Path ${cached})) )  {
        Remove-Variable cached
    }

    If ( ${ID} ) {
        $R = Get-OpenFile -Id ${ID}
    } 
    ElseIf ( ${filePath} -and ${cached} ) {
        $R = Get-OpenFile -filter "${filePath}" -cached "${cached}" #| Format-Table -Property ID,FilePath,ProcessName -auto -wrap
    } 
    Else {
        $R = Get-OpenFile -filter "${FilePath}"
    }

    If ($R.Count -eq 1) {
        Remove-OpenFile -id $($R.ID)
    } 
    Else {
        Write-Error "No valid file handle found. Handle ${ID} appears $($R.Count) times."
    }

}; Set-Alias delflock Remove-TempFile -Description "Displays X."

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

}; Set-Alias trackflock Set-OpenFile -Description "Controls the 'maintain objects list' so openfile can track."


Function Get-LockedFile { # Displays applications with opened or locked files in local file system.
    [CmdletBinding(DefaultParametersetName="Standard")]
    Param(
        [Parameter(
            ParameterSetName='Standard',
            Position=0,
            ValueFromPipeline=$True
        )]
        [STRING]$Filter,
        [Parameter(ParameterSetName='Standard',Position=1)]
        [STRING]$Cached,
        [Parameter(ParameterSetName='Standard',Position=0)]
        [INT]$ID
    )
  
    Begin{

        ${FileCache}= ${Cached};
        ${StartStr} = "Files Opened Locally:"
        ${EndStr}   = "Files opened remotely via local share points:"
        ${Header}   = "ID","AccessedBy","PID","ProcessName","FilePath"
        ${Pattern}  = $Filter 

    }
    Process{

        Try {

            If ( ${FileCache} -and (Test-Path ${FileCache}) ) {                

                ${FileObj} = Get-Item ${FileCache};
                ${Message} = "Using cached data in ${FileCache} generated $(${FileObj}.LastWriteTime)";
                ${TextRaw} = Get-Content ${FileCache};

            } Else {

                Write-Output "No cache found at ${FileCache}, so scanning operation may take several minutes to complete. `nPlease do not close this session. "
                $TimeResult = Measure-Command -Expression { 
                    ${TextRaw} = openfiles /query /v /fo CSV /nh;
                } 
                If (${FileCache}) { ${TextRaw} > ${FileCache} }
                ${Message} = "Openfiles completed in $($TimeResult.TotalMinutes) minutes."

            }

            ${OpenFiles} = ${TextRaw}[($TextRaw.IndexOf(${StartStr}) + 2) .. ($TextRaw.IndexOf(${EndStr}) -3)] | ConvertFrom-Csv -Header ${Header};

        } 
        Catch {
            Write-Output $Error[0]
        }
        Finally{
            Write-Verbose ${Message}
        }

    }
    End{
        
        If ( $ID ) {
            ${OpenFiles} |
                Where-Object { $_.ID -eq "${ID}" } 
        } ElseIf ( $Filter ) {
            $FinalResult = ${OpenFiles} |
                Where-Object { $_.FilePath -Like "*${Pattern}*" }
        } Else {
            $FinalResult = ${OpenFiles} 
        }

        $FinalResult | Format-Table -Property ID,FilePath,PID,ProcessName -auto -wrap

    }
}; Set-Alias flock Get-LockedFile -Description "Displays applications with opened or locked files in local file system."


Function Get-Windows { # Identifies Windows Operating System: Maj.Min.Build.Release.Update

    [CmdletBinding(DefaultParameterSetName="Version")]
    Param(
        [Parameter(ParameterSetName='MajorNo',Position=0,Mandatory=$False)][SWITCH]$Major,
        [Parameter(ParameterSetName='MinorNo',Position=0,Mandatory=$False)][SWITCH]$Minor,
        [Parameter(ParameterSetName='BuildNo',Position=0,Mandatory=$False)][SWITCH]$Build,
        [Parameter(ParameterSetName='RevisionNo',Position=0,Mandatory=$False)][SWITCH]$Revision,
        [Parameter(ParameterSetName='ReleaseID',Position=0,Mandatory=$False)][SWITCH]$Release,
        [Parameter(ParameterSetName='Role',Position=0,Mandatory=$False)][SWITCH]$Role,
        [Parameter(ParameterSetName='Product',Position=0,Mandatory=$False)][SWITCH]$Product,
        [Parameter(ParameterSetName='Version',Position=0,Mandatory=$False)][SWITCH]$Version,
        [Parameter(ParameterSetName='Serial',Position=0,Mandatory=$False)][SWITCH]$SerialNumber,
        [Parameter(ParameterSetName='Up',Position=0,Mandatory=$False)][SWITCH]$Uptime
    )

    Begin {
        
        $w32OS = Get-CimInstance -ClassName win32_operatingsystem -EA 0
        
        $Ver = $w32OS.Version.ToString().Split(".");
        $UBR = [STRING](Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' UBR).UBR
        $objVersion = New-Object -typename System.Version $Ver[0],$Ver[1],$Ver[2],$UBR;
        $strVersion = $w32OS.Version + "." + $UBR

        $DomainRole = @{
            0 = "Standalone Workstation" ;
            1 = "Member Workstation" ;
            2 = "Standalone Server" ;
            3 = "Member Server" ;
            4 = "Backup Domain Controller" ;
            5 = "Primary Domain Controller" 
        } 

    }
    Process{
        
        switch($PSCmdlet.ParameterSetName){
            
            "MajorNo"    { $objVersion.Major.ToString(); Break; }

            "MinorNo"    { $objVersion.Minor.ToString(); Break; }

            "BuildNo"    { $objVersion.Build.ToString(); Break; }

            "Product"    { $w32OS.Caption; Break; }

            "RevisionNo" { $objVersion.Revision.ToString(); Break; }

            "ReleaseID"  { (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId).ReleaseId
                           Break; 
                         }

            "Role"       { $w32CS = Get-CimInstance -ClassName win32_computersystem -EA 0; 
                           $DomainRole[[int]$w32CS.DomainRole]; 
                           Break; 
                         }

            "Serial"     { $w32OS.SerialNumber; Break; }

            "Up"         { $when = $w32OS.LastBootUpTime.GetDateTimeFormats()[112];
                           $T = (Get-Date).Subtract($w32OS.LastBootUpTime);
                           $U = $T.Subtract($T.Days); 
                           " $When up $($T.Days) days, $($U.Hours)`:$($U.Minutes)"; 
                           Break; 
                         }

            "Version"    { $strVersion; Break; }
            
            Default      { $objVersion }

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
        Get-WinVersion
        On Windows 10 released April 2018, this command returned "10.0.17134.228"

    #>

}; Set-Alias version Get-Windows -Description "Identifies Windows Operating System: Maj.Min.Build.Release.Update."



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