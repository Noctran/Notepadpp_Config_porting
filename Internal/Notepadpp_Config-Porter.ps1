#requires -version 5

#region Preparations
#region Header
<#
.SYNOPSIS
  Exports and Imports Notepad++ Configuration
.DESCRIPTION
  <Brief description of script>
.PARAMETER Import
    Directs the Script to copy Config Files from Tool internal Storage to Notepad++ Directories.
.PARAMETER Export
    Directs the Script to copy Config Files from Notepad++ Directories to Tool internal Storage.
.PARAMETER IgnoreNotepadppInstalled
    Ignores Checks if Notepad++ is installed
.PARAMETER InstallWithWinget
    Installs Notepad++ via Winget
#>

<# Script Version Infos
  Current Info:
    Version:        1.0
      Author:         Noctran
      Creation Date:  17.07.2022
      Purpose/Change: Initial development

  Version History:

#>#Script Version Infos

#endregion Header
#region -----------------------------------------------------------[Parameters]-----------------------------------------------------------

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [parameter(mandatory = $true, ParameterSetName = 'Import')][switch]$Import,
    [parameter(mandatory = $true, ParameterSetName = 'Export')][switch]$Export,

    [parameter(mandatory = $false)][switch]$IgnoreNotepadppInstalled,
    [parameter(mandatory = $false)][switch]$InstallWithWinget
)

#endregion Parameters
#region Initial Tasks
#region -----------------------------------------------------------[Initialisation]-----------------------------------------------------------

# Loading required Modules

#region Common Variables
$VerbosePreference = 'Continue'
$WarningPreference = 'Continue' #'Inquire'
$DebugPreference = 'SilentlyContinue' #'Continue', 'Inquire'
$ErrorActionPreference = , 'Continue' #'SilentlyContinue', 'Inquire'
$InformationPreference = 'Continue'
$ErrorView = 'CategoryView'
#$WhatIfPreference = $true
#endregion Common Variables

#endregion Initialisation
#region ----------------------------------------------------------[Declarations]----------------------------------------------------------

#region Development Related Variables
[string]$ScriptName = $MyInvocation.MyCommand.Name
[system.version]$NocScriptVersion = '1.0.0.0'
[datetime]$NocScriptDate = '2022.07.17'
[string]$NocScriptAuthor = 'Noctran'
#endregion Development Related Variables

$NocSystemAppdata = Join-Path -Path "$env:APPDATA" -ChildPath 'Notepad++'
$NocSystemProgramFiles = Join-Path -Path "$env:ProgramFiles" -ChildPath 'Notepad++'
$NocToolPath = (Get-Item -Path $PSScriptRoot ).parent.FullName
$NocToolParentPath = (Get-Item -Path $PSScriptRoot ).parent.parent.FullName
$NocToolAppdataStorage = Join-Path -Path "$NocToolPath" -ChildPath 'Storage\ToolAppdata'
$NocToolProgramFilesStorage = Join-Path -Path "$NocToolPath" -ChildPath 'Storage\ToolProgramFiles'

#endregion Declarations
#region -----------------------------------------------------------[Functions]------------------------------------------------------------

Function Write-NocSLog {
    <#
    .SYNOPSIS
      Standalone Simple Logging
    .DESCRIPTION
      <Brief description of Function>
    .PARAMETER logLevel
      Valide Inputs: 0-3
        0 { $textColor = "White"}
        1 { $textColor = "Green"}
        2 { $textColor = "Yellow"}
        3 { $textColor = "Red"}
    .PARAMETER logLine
      Log Message
    .OUTPUTS
      Writes Log Message back into Shell
    .EXAMPLE
      Write-NocSLog 2 "Example"
    #>

    <# Script Version Infos

    Current Info:

      Version:        0.1
        Author:         Noctran
        Creation Date:  09.06.2022
        Purpose/Change: Initial Function development. Bare Minimum, only Logging to Shell.

    #>

    #region -----------------------------------------------------------[Parameters]-----------------------------------------------------------

    #script requires variable $log set
    #logType Currently not in use, default(1) only one enabled.
    #logType = 0 Print & Log, 1 = Print Only, 2 = LogOnly
    Param (
        [Parameter(Mandatory = $true, Position = 0)][ValidatePattern('[0-3]')][int]$logLevel,
        [Parameter(Mandatory = $true, Position = 1)][AllowEmptyString()][String]$logMessage,
        [Parameter(Mandatory = $false)][switch]$printNoNewLine,
        [Parameter(Mandatory = $false)][switch]$SectionHeader,
        [Parameter(Mandatory = $false)][switch]$SectionFooter
    )

    #endregion Parameters
    #region -----------------------------------------------------------[Execution]------------------------------------------------------------

    Begin {
        #color & Type
        switch ($logLevel) {
            0 { $textColor = 'White'}
            1 { $textColor = 'Green'}
            2 { $textColor = 'Yellow'}
            3 { $textColor = 'Red'}
        }
    }
    Process {
        #print based on type

        if ($SectionHeader) {
            Write-Host ''
            Write-Host ''
        }
        if ($SectionFooter) {
            Write-Host ''
        }

        Write-Host $logMessage -ForegroundColor $textColor -NoNewline

        if ($SectionHeader) {
            Write-Host ''
        }
        if ($SectionFooter) {
            Write-Host ''
            Write-Host ''
        }
    }
    End {
        #create newline after text printing if printnonewline not specified
        if (-not($printNoNewLine) -and ($logType -ne 2)) {
            Write-Host ''
        }
    }

    #endregion Execution
    #End of Function
}

Function Test-DirectoryEmtpy {
    param (
        [parameter(mandatory = $true)][string]$Directory
    )
    $directoryInfo = Get-ChildItem $Directory | Measure-Object
    if ($directoryInfo.count -eq 0) {$true} else {$false}
}

Function New-ToolStorageDir {
    param (
        [parameter(mandatory = $true)][string]$Directory
    )
    if (Test-Path -Path $Directory) {
        if (Test-DirectoryEmtpy $Directory) {
                Get-ChildItem -Path $Directory | Remove-Item -Recurse -Force
            }
    } else {
        Write-NocSLog 0 "Directory `"$Directory`" does not Exist, Creating it now"
        New-Item -Path $Directory -ItemType Directory
    }
}

Function Test-ToolStorageDir {
    param (
        [parameter(mandatory = $true)][string]$Directory
    )
    if (Test-Path -Path $Directory) {
        if (Test-DirectoryEmtpy $Directory) {
            Write-NocSLog 2 "Directory: `"$Directory`"  is empty"
        }
    } else {
        Write-NocSLog 3 "Directory `"$Directory`" does not Exist,"
        $NocAbortToolExecution = $true
        break
    }
}

Function copy-SubDir {
    param (
        [parameter(mandatory = $true)][string]$SubDirectory
    )

    Write-NocSLog 0 " Subdir: `"$SubDirectory`" "

    $ProgramFilesSystemPath = Join-Path -Path $NocSystemProgramFiles -ChildPath $SubDirectory
    $ProgramFilesStoragePath = Join-Path -Path $NocToolProgramFilesStorage -ChildPath $SubDirectory

    Write-NocSLog 0 "Creating Directory `"$ProgramFilesStoragePath`""
    New-Item -Path "$ProgramFilesStoragePath" -ItemType Directory

    Get-ChildItem -Path "$ProgramFilesSystemPath" |
        Copy-Item -Destination "$ProgramFilesStoragePath\" -Force -Recurse
}

Function stop-Notepadpp {
    taskkill /IM 'notepad++.exe'/F
}

#endregion Functions
#endregion Initial Tasks
#endregion Preparations
#region Main Script Execution
#region -----------------------------------------------------------[Execution]------------------------------------------------------------

Write-NocSLog 0 "Executing Script: `"$ScriptName`" in Version: `"$NocScriptVersion`""
stop-Notepadpp

#region Notepad Check
if ($NocAbortToolExecution) {Write-NocSLog 3 'Aborting Script Execution', exit 1}

$w32 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName -Like 'NotePad++*'
$w64 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName -Like 'NotePad++*'
if ($w64 -or $w32) {
    Write-NocSLog 0 'Notepad++ is installed, execution Tool'
} elseif ($InstallWithWinget) {
    winget install 'notepad++.notepad++'
} Else {
    if ($IgnoreNotepadppInstalled) {
        Write-NocSLog 2 'Notepad++ is not installed on this Device, execution still due tue given Parameter'
    } else {
        Write-NocSLog 3 'Notepad++ is not installed on this Device, aborting execution'
        $NocAbortToolExecution = $true
    }
}
#endregion Notepad Check

#region Tool
if ($NocAbortToolExecution) {Write-NocSLog 3 'Aborting Script Execution', exit 1}

if ($export) {
    New-ToolStorageDir $NocToolAppdataStorage
    New-ToolStorageDir $NocToolProgramFilesStorage
}

if ($import) {
    Test-ToolStorageDir $NocToolAppdataStorage
    Test-ToolStorageDir $NocToolProgramFilesStorage
}
#endregion Tool

#region Config
if ($NocAbortToolExecution) {Write-NocSLog 3 'Aborting Script Execution', exit 1}

#region Exporting Config
if ($export) {
    if ($NocAbortToolExecution) {Write-NocSLog 3 'Aborting Script Execution', exit 1}
    #TODO Add Ability to get Files from Remote Machine

    Copy-Item -Path "$NocSystemAppdata\*" -Destination "$NocToolAppdataStorage" -Recurse -Force

    copy-SubDir plugins
    copy-SubDir localization
    copy-SubDir functionList
    copy-SubDir autoCompletion
    Copy-Item -Path "$NocSystemProgramFiles\*" -Destination "$NocToolProgramFilesStorage" -Include '*.xml'
}
#endregion Exporting Config

#region Importing Config
if ($import) {
    if ($NocAbortToolExecution) {Write-NocSLog 3 'Aborting Script Execution', exit 1}

    #TODO Check if there is already a custom Config
    #TODO Ability to merge existing Configs

    Copy-Item -Path "$NocToolAppdataStorage\*" -Destination "$NocSystemAppdata" -Recurse -Force
    Copy-Item -Path "$NocToolProgramFilesStorage\*" -Destination "$NocSystemProgramFiles" -Recurse -Force

}
#endregion Importing Config

#endregion Config

#endregion Execution
#region -----------------------------------------------------------[Finish up]------------------------------------------------------------

if ($export) {
    Copy-Item -Path "$NocToolPath\Internal\Import.cmd" -Destination "$NocToolPath" -Force
    Copy-Item -Path "$NocToolPath\Internal\InstallViaWingetAndImport.cmd" -Destination "$NocToolPath" -Force
    Remove-Item -Path "$NocToolPath\Export.cmd"

    Compress-Archive -Path "$NocToolPath\*" -DestinationPath "$NocToolParentPath\Notepadpp_Config_Export.zip"-Force -CompressionLevel 'Fastest'
}

if ($import) {
    Start-Process notepad++
}

#Cleanup Actions
# Error Handling
Trap {
    $TerminatingError = $_
    Write-NocSLog 3 "A not solvable terminating error occurred, ending the  Script: `"$ScriptName`""
    Write-NocSLog 3 "$TerminatingError"
    break
}

Write-NocSLog 0 "Finished Executing Script: `"$ScriptName`" in Version: `"$ACPScriptVersion`""

#endregion Finish up
#endregion Main Script Execution
#End of Script
