#region Function Write-FunctionHeaderOrFooter (From PowerShell App Deploy Toolkit)
Function Write-FunctionHeaderOrFooter
{
    <#
    .SYNOPSIS
        Write the function header or footer to the log upon first entering or exiting a function.
    .DESCRIPTION
        Write the "Function Start" message, the bound parameters the function was invoked with, or the "Function End" message when entering or exiting a function.
        Messages are debug messages so will only be logged if LogDebugMessage option is enabled in XML config file.
    .PARAMETER CmdletName
        The name of the function this function is invoked from.
    .PARAMETER CmdletBoundParameters
        The bound parameters of the function this function is invoked from.
    .PARAMETER Header
        Write the function header.
    .PARAMETER Footer
        Write the function footer.
    .EXAMPLE
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    .EXAMPLE
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    .NOTES
        This is an internal script function and should typically not be called directly.
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$CmdletName,
        [Parameter(Mandatory = $true, ParameterSetName = 'Header')]
        [AllowEmptyCollection()]
        [hashtable]$CmdletBoundParameters,
        [Parameter(Mandatory = $true, ParameterSetName = 'Header')]
        [switch]$Header,
        [Parameter(Mandatory = $true, ParameterSetName = 'Footer')]
        [switch]$Footer
    )
        
    If ($Header)
    {
        Write-Log -Message 'Function Start' -Source ${CmdletName} -DebugMessage
            
        ## Get the parameters that the calling function was invoked with
        [string]$CmdletBoundParameters = $CmdletBoundParameters | Format-Table -Property @{ Label = 'Parameter'; Expression = { "[-$($_.Key)]" } }, @{ Label = 'Value'; Expression = { $_.Value }; Alignment = 'Left' }, @{ Label = 'Type'; Expression = { $_.Value.GetType().Name }; Alignment = 'Left' } -AutoSize -Wrap | Out-String
        If ($CmdletBoundParameters)
        {
            Write-Log -Message "Function invoked with bound parameter(s): `n$CmdletBoundParameters" -Source ${CmdletName} -DebugMessage
        }
        Else
        {
            Write-Log -Message 'Function invoked without any bound parameters.' -Source ${CmdletName} -DebugMessage
        }
    }
    ElseIf ($Footer)
    {
        Write-Log -Message 'Function End' -Source ${CmdletName} -DebugMessage
    }
}
#endregion

#region Function Write-Log (From PowerShell App Deploy Toolkit)
Function Write-Log
{
    <#
    .SYNOPSIS
        Write messages to a log file in CMTrace.exe compatible format or Legacy text file format.
    .DESCRIPTION
        Write messages to a log file in CMTrace.exe compatible format or Legacy text file format and optionally display in the console.
    .PARAMETER Message
        The message to write to the log file or output to the console.
    .PARAMETER Severity
        Defines message type. When writing to console or CMTrace.exe log format, it allows highlighting of message type.
        Options: 1 = Information (default), 2 = Warning (highlighted in yellow), 3 = Error (highlighted in red)
    .PARAMETER Source
        The source of the message being logged.
    .PARAMETER ScriptSection
        The heading for the portion of the script that is being executed. Default is: $script:installPhase.
    .PARAMETER LogType
        Choose whether to write a CMTrace.exe compatible log file or a Legacy text log file.
    .PARAMETER LogFileDirectory
        Set the directory where the log file will be saved.
    .PARAMETER LogFileName
        Set the name of the log file.
    .PARAMETER MaxLogFileSizeMB
        Maximum file size limit for log file in megabytes (MB). Default is 10 MB.
    .PARAMETER WriteHost
        Write the log message to the console.
    .PARAMETER ContinueOnError
        Suppress writing log message to console on failure to write message to log file. Default is: $true.
    .PARAMETER PassThru
        Return the message that was passed to the function
    .PARAMETER DebugMessage
        Specifies that the message is a debug message. Debug messages only get logged if -LogDebugMessage is set to $true.
    .PARAMETER LogDebugMessage
        Debug messages only get logged if this parameter is set to $true in the config XML file.
    .EXAMPLE
        Write-Log -Message "Installing patch MS15-031" -Source 'Add-Patch' -LogType 'CMTrace'
    .EXAMPLE
        Write-Log -Message "Script is running on Windows 8" -Source 'Test-ValidOS' -LogType 'Legacy'
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [Alias('Text')]
        [string[]]$Message,
        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateRange(1, 3)]
        [int16]$Severity = 1,
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNull()]
        [string]$Source = '',
        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateNotNullorEmpty()]
        [string]$ScriptSection = $script:installPhase,
        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateSet('CMTrace', 'Legacy')]
        [string]$LogType = $configToolkitLogStyle,
        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateNotNullorEmpty()]
        [string]$LogFileDirectory = $(If ($configToolkitCompressLogs) { $logTempFolder } Else { $configToolkitLogDir }),
        [Parameter(Mandatory = $false, Position = 6)]
        [ValidateNotNullorEmpty()]
        [string]$LogFileName = $logName,
        [Parameter(Mandatory = $false, Position = 7)]
        [ValidateNotNullorEmpty()]
        [decimal]$MaxLogFileSizeMB = $configToolkitLogMaxSize,
        [Parameter(Mandatory = $false, Position = 8)]
        [ValidateNotNullorEmpty()]
        [boolean]$WriteHost = $configToolkitLogWriteToHost,
        [Parameter(Mandatory = $false, Position = 9)]
        [ValidateNotNullorEmpty()]
        [boolean]$ContinueOnError = $true,
        [Parameter(Mandatory = $false, Position = 10)]
        [switch]$PassThru = $false,
        [Parameter(Mandatory = $false, Position = 11)]
        [switch]$DebugMessage = $false,
        [Parameter(Mandatory = $false, Position = 12)]
        [boolean]$LogDebugMessage = $configToolkitLogDebugMessage
    )
        
    Begin
    {
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            
        ## Logging Variables
        #  Log file date/time
        [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
        [string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
        If (-not (Test-Path -LiteralPath 'variable:LogTimeZoneBias')) { [int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes }
        [string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
        #  Initialize variables
        [boolean]$ExitLoggingFunction = $false
        If (-not (Test-Path -LiteralPath 'variable:DisableLogging')) { $DisableLogging = $false }
        #  Check if the script section is defined
        [boolean]$ScriptSectionDefined = [boolean](-not [string]::IsNullOrEmpty($ScriptSection))
        #  Get the file name of the source script
        Try
        {
            If ($script:MyInvocation.Value.ScriptName)
            {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
            }
            Else
            {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
            }
        }
        Catch
        {
            $ScriptSource = ''
        }
            
        ## Create script block for generating CMTrace.exe compatible log entry
        [scriptblock]$CMTraceLogString = {
            Param (
                [string]$lMessage,
                [string]$lSource,
                [int16]$lSeverity
            )
            "<![LOG[$lMessage]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$lSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$lSeverity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"
        }
            
        ## Create script block for writing log entry to the console
        [scriptblock]$WriteLogLineToHost = {
            Param (
                [string]$lTextLogLine,
                [int16]$lSeverity
            )
            If ($WriteHost)
            {
                #  Only output using color options if running in a host which supports colors.
                If ($Host.UI.RawUI.ForegroundColor)
                {
                    Switch ($lSeverity)
                    {
                        3 { Write-Host -Object $lTextLogLine -ForegroundColor 'Red' -BackgroundColor 'Black' }
                        2 { Write-Host -Object $lTextLogLine -ForegroundColor 'Yellow' -BackgroundColor 'Black' }
                        1 { Write-Host -Object $lTextLogLine }
                    }
                }
                #  If executing "powershell.exe -File <filename>.ps1 > log.txt", then all the Write-Host calls are converted to Write-Output calls so that they are included in the text log.
                Else
                {
                    Write-Output -InputObject $lTextLogLine
                }
            }
        }
            
        ## Exit function if it is a debug message and logging debug messages is not enabled in the config XML file
        If (($DebugMessage) -and (-not $LogDebugMessage)) { [boolean]$ExitLoggingFunction = $true; Return }
        ## Exit function if logging to file is disabled and logging to console host is disabled
        If (($DisableLogging) -and (-not $WriteHost)) { [boolean]$ExitLoggingFunction = $true; Return }
        ## Exit Begin block if logging is disabled
        If ($DisableLogging) { Return }
        ## Exit function function if it is an [Initialization] message and the toolkit has been relaunched
        If (($AsyncToolkitLaunch) -and ($ScriptSection -eq 'Initialization')) { [boolean]$ExitLoggingFunction = $true; Return }
            
        ## Create the directory where the log file will be saved
        If (-not (Test-Path -LiteralPath $LogFileDirectory -PathType 'Container'))
        {
            Try
            {
                $null = New-Item -Path $LogFileDirectory -Type 'Directory' -Force -ErrorAction 'Stop'
            }
            Catch
            {
                [boolean]$ExitLoggingFunction = $true
                #  If error creating directory, write message to console
                If (-not $ContinueOnError)
                {
                    Write-Host -Object "[$LogDate $LogTime] [${CmdletName}] $ScriptSection :: Failed to create the log directory [$LogFileDirectory]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                }
                Return
            }
        }
            
        ## Assemble the fully qualified path to the log file
        [string]$LogFilePath = Join-Path -Path $LogFileDirectory -ChildPath $LogFileName
    }
    Process
    {
        ## Exit function if logging is disabled
        If ($ExitLoggingFunction) { Return }
            
        ForEach ($Msg in $Message)
        {
            ## If the message is not $null or empty, create the log entry for the different logging methods
            [string]$CMTraceMsg = ''
            [string]$ConsoleLogLine = ''
            [string]$LegacyTextLogLine = ''
            If ($Msg)
            {
                #  Create the CMTrace log message
                If ($ScriptSectionDefined) { [string]$CMTraceMsg = "[$ScriptSection] :: $Msg" }
                    
                #  Create a Console and Legacy "text" log entry
                [string]$LegacyMsg = "[$LogDate $LogTime]"
                If ($ScriptSectionDefined) { [string]$LegacyMsg += " [$ScriptSection]" }
                If ($Source)
                {
                    [string]$ConsoleLogLine = "$LegacyMsg [$Source] :: $Msg"
                    Switch ($Severity)
                    {
                        3 { [string]$LegacyTextLogLine = "$LegacyMsg [$Source] [Error] :: $Msg" }
                        2 { [string]$LegacyTextLogLine = "$LegacyMsg [$Source] [Warning] :: $Msg" }
                        1 { [string]$LegacyTextLogLine = "$LegacyMsg [$Source] [Info] :: $Msg" }
                    }
                }
                Else
                {
                    [string]$ConsoleLogLine = "$LegacyMsg :: $Msg"
                    Switch ($Severity)
                    {
                        3 { [string]$LegacyTextLogLine = "$LegacyMsg [Error] :: $Msg" }
                        2 { [string]$LegacyTextLogLine = "$LegacyMsg [Warning] :: $Msg" }
                        1 { [string]$LegacyTextLogLine = "$LegacyMsg [Info] :: $Msg" }
                    }
                }
            }
                
            ## Execute script block to create the CMTrace.exe compatible log entry
            [string]$CMTraceLogLine = & $CMTraceLogString -lMessage $CMTraceMsg -lSource $Source -lSeverity $Severity
                
            ## Choose which log type to write to file
            If ($LogType -ieq 'CMTrace')
            {
                [string]$LogLine = $CMTraceLogLine
            }
            Else
            {
                [string]$LogLine = $LegacyTextLogLine
            }
                
            ## Write the log entry to the log file if logging is not currently disabled
            If (-not $DisableLogging)
            {
                Try
                {
                    $LogLine | Out-File -FilePath $LogFilePath -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Stop'
                }
                Catch
                {
                    If (-not $ContinueOnError)
                    {
                        Write-Host -Object "[$LogDate $LogTime] [$ScriptSection] [${CmdletName}] :: Failed to write message [$Msg] to the log file [$LogFilePath]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                    }
                }
            }
                
            ## Execute script block to write the log entry to the console if $WriteHost is $true
            & $WriteLogLineToHost -lTextLogLine $ConsoleLogLine -lSeverity $Severity
        }
    }
    End
    {
        ## Archive log file if size is greater than $MaxLogFileSizeMB and $MaxLogFileSizeMB > 0
        Try
        {
            If ((-not $ExitLoggingFunction) -and (-not $DisableLogging))
            {
                [IO.FileInfo]$LogFile = Get-ChildItem -LiteralPath $LogFilePath -ErrorAction 'Stop'
                [decimal]$LogFileSizeMB = $LogFile.Length / 1MB
                If (($LogFileSizeMB -gt $MaxLogFileSizeMB) -and ($MaxLogFileSizeMB -gt 0))
                {
                    ## Change the file extension to "lo_"
                    [string]$ArchivedOutLogFile = [IO.Path]::ChangeExtension($LogFilePath, 'lo_')
                    [hashtable]$ArchiveLogParams = @{ ScriptSection = $ScriptSection; Source = ${CmdletName}; Severity = 2; LogFileDirectory = $LogFileDirectory; LogFileName = $LogFileName; LogType = $LogType; MaxLogFileSizeMB = 0; WriteHost = $WriteHost; ContinueOnError = $ContinueOnError; PassThru = $false }
                        
                    ## Log message about archiving the log file
                    $ArchiveLogMessage = "Maximum log file size [$MaxLogFileSizeMB MB] reached. Rename log file to [$ArchivedOutLogFile]."
                    Write-Log -Message $ArchiveLogMessage @ArchiveLogParams
                        
                    ## Archive existing log file from <filename>.log to <filename>.lo_. Overwrites any existing <filename>.lo_ file. This is the same method SCCM uses for log files.
                    Move-Item -LiteralPath $LogFilePath -Destination $ArchivedOutLogFile -Force -ErrorAction 'Stop'
                        
                    ## Start new log file and Log message about archiving the old log file
                    $NewLogMessage = "Previous log file was renamed to [$ArchivedOutLogFile] because maximum log file size of [$MaxLogFileSizeMB MB] was reached."
                    Write-Log -Message $NewLogMessage @ArchiveLogParams
                }
            }
        }
        Catch
        {
            ## If renaming of file fails, script will continue writing to log file even if size goes over the max file size
        }
        Finally
        {
            If ($PassThru) { Write-Output -InputObject $Message }
        }
    }
}
#endregion

#region Function New-Folder (From PowerShell App Deploy Toolkit)
Function New-Folder
{
    <#
    .SYNOPSIS
        Create a new folder.
    .DESCRIPTION
        Create a new folder if it does not exist.
    .PARAMETER Path
        Path to the new folder to create.
    .PARAMETER ContinueOnError
        Continue if an error is encountered. Default is: $true.
    .EXAMPLE
        New-Folder -Path "$envWinDir\System32"
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [boolean]$ContinueOnError = $true
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            If (-not (Test-Path -LiteralPath $Path -PathType 'Container'))
            {
                Write-Log -Message "Create folder [$Path]." -Source ${CmdletName}
                $null = New-Item -Path $Path -ItemType 'Directory' -ErrorAction 'Stop'
            }
            Else
            {
                Write-Log -Message "Folder [$Path] already exists." -Source ${CmdletName}
            }
        }
        Catch
        {
            Write-Log -Message "Failed to create folder [$Path]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            If (-not $ContinueOnError)
            {
                Throw "Failed to create folder [$Path]: $($_.Exception.Message)"
            }
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Copy-File (From PowerShell App Deploy Toolkit)
Function Copy-File
{
    <#
    .SYNOPSIS
        Copy a file or group of files to a destination path.
    .DESCRIPTION
        Copy a file or group of files to a destination path.
    .PARAMETER Path
        Path of the file to copy.
    .PARAMETER Destination
        Destination Path of the file to copy.
    .PARAMETER Recurse
        Copy files in subdirectories.
    .PARAMETER Flatten
        Flattens the files into the root destination directory.
    .PARAMETER ContinueOnError
        Continue if an error is encountered. This will continue the deployment script, but will not continue copying files if an error is encountered. Default is: $true.
    .PARAMETER ContinueFileCopyOnError
        Continue copying files if an error is encountered. This will continue the deployment script and will warn about files that failed to be copied. Default is: $false.
    .EXAMPLE
        Copy-File -Path "$dirSupportFiles\MyApp.ini" -Destination "$envWindir\MyApp.ini"
    .EXAMPLE
        Copy-File -Path "$dirSupportFiles\*.*" -Destination "$envTemp\tempfiles"
        Copy all of the files in a folder to a destination folder.
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string[]]$Path,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Destination,
        [Parameter(Mandatory = $false)]
        [switch]$Recurse = $false,
        [Parameter(Mandatory = $false)]
        [switch]$Flatten,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [boolean]$ContinueOnError = $true,
        [ValidateNotNullOrEmpty()]
        [boolean]$ContinueFileCopyOnError = $false
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            $null = $fileCopyError
            If ((-not ([IO.Path]::HasExtension($Destination))) -and (-not (Test-Path -LiteralPath $Destination -PathType 'Container')))
            {
                Write-Log -Message "Destination folder does not exist, creating destination folder [$destination]." -Source ${CmdletName}
                $null = New-Item -Path $Destination -Type 'Directory' -Force -ErrorAction 'Stop'
            }
    
            if ($Flatten)
            {
                If ($Recurse)
                {
                    Write-Log -Message "Copy file(s) recursively in path [$path] to destination [$destination] root folder, flattened." -Source ${CmdletName}
                    If (-not $ContinueFileCopyOnError)
                    {
                        $null = Get-ChildItem -Path $path -Recurse | Where-Object { !($_.PSIsContainer) } | ForEach {
                            Copy-Item -Path ($_.FullName) -Destination $destination -Force -ErrorAction 'Stop'
                        }
                    }
                    Else
                    {
                        $null = Get-ChildItem -Path $path -Recurse | Where-Object { !($_.PSIsContainer) } | ForEach {
                            Copy-Item -Path ($_.FullName) -Destination $destination -Force -ErrorAction 'SilentlyContinue' -ErrorVariable FileCopyError
                        }
                    }
                }
                Else
                {
                    Write-Log -Message "Copy file in path [$path] to destination [$destination]." -Source ${CmdletName}
                    If (-not $ContinueFileCopyOnError)
                    {
                        $null = Copy-Item -Path $path -Destination $destination -Force -ErrorAction 'Stop'
                    }
                    Else
                    {
                        $null = Copy-Item -Path $path -Destination $destination -Force -ErrorAction 'SilentlyContinue' -ErrorVariable FileCopyError
                    }
                }
            }
            Else
            {
                $null = $FileCopyError
                If ($Recurse)
                {
                    Write-Log -Message "Copy file(s) recursively in path [$path] to destination [$destination]." -Source ${CmdletName}
                    If (-not $ContinueFileCopyOnError)
                    {
                        $null = Copy-Item -Path $Path -Destination $Destination -Force -Recurse -ErrorAction 'Stop'
                    }
                    Else
                    {
                        $null = Copy-Item -Path $Path -Destination $Destination -Force -Recurse -ErrorAction 'SilentlyContinue' -ErrorVariable FileCopyError
                    }
                }
                Else
                {
                    Write-Log -Message "Copy file in path [$path] to destination [$destination]." -Source ${CmdletName}
                    If (-not $ContinueFileCopyOnError)
                    {
                        $null = Copy-Item -Path $Path -Destination $Destination -Force -ErrorAction 'Stop'
                    }
                    Else
                    {
                        $null = Copy-Item -Path $Path -Destination $Destination -Force -ErrorAction 'SilentlyContinue' -ErrorVariable FileCopyError
                    }
                }
            }
                
            If ($fileCopyError)
            { 
                Write-Log -Message "The following warnings were detected while copying file(s) in path [$path] to destination [$destination]. `n$FileCopyError" -Severity 2 -Source ${CmdletName}
            }
            Else
            {
                Write-Log -Message "File copy completed successfully." -Source ${CmdletName}            
            }
        }
        Catch
        {
            Write-Log -Message "Failed to copy file(s) in path [$path] to destination [$destination]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            If (-not $ContinueOnError)
            {
                Throw "Failed to copy file(s) in path [$path] to destination [$destination]: $($_.Exception.Message)"
            }
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Convert-RegistryPath (From PowerShell App Deploy Toolkit)
Function Convert-RegistryPath
{
    <#
    .SYNOPSIS
        Converts the specified registry key path to a format that is compatible with built-in PowerShell cmdlets.
    .DESCRIPTION
        Converts the specified registry key path to a format that is compatible with built-in PowerShell cmdlets.
        Converts registry key hives to their full paths. Example: HKLM is converted to "Registry::HKEY_LOCAL_MACHINE".
    .PARAMETER Key
        Path to the registry key to convert (can be a registry hive or fully qualified path)
    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.
        Specify this parameter from the Invoke-HKCURegistrySettingsForAllUsers function to read/edit HKCU registry settings for all users on the system.
    .EXAMPLE
        Convert-RegistryPath -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1AD147D0-BE0E-3D6C-AC11-64F6DC4163F1}'
    .EXAMPLE
        Convert-RegistryPath -Key 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1AD147D0-BE0E-3D6C-AC11-64F6DC4163F1}'
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Key,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [string]$SID
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        ## Convert the registry key hive to the full path, only match if at the beginning of the line
        If ($Key -match '^HKLM:\\|^HKCU:\\|^HKCR:\\|^HKU:\\|^HKCC:\\|^HKPD:\\')
        {
            #  Converts registry paths that start with, e.g.: HKLM:\
            $key = $key -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'
            $key = $key -replace '^HKCR:\\', 'HKEY_CLASSES_ROOT\'
            $key = $key -replace '^HKCU:\\', 'HKEY_CURRENT_USER\'
            $key = $key -replace '^HKU:\\', 'HKEY_USERS\'
            $key = $key -replace '^HKCC:\\', 'HKEY_CURRENT_CONFIG\'
            $key = $key -replace '^HKPD:\\', 'HKEY_PERFORMANCE_DATA\'
        }
        ElseIf ($Key -match '^HKLM:|^HKCU:|^HKCR:|^HKU:|^HKCC:|^HKPD:')
        {
            #  Converts registry paths that start with, e.g.: HKLM:
            $key = $key -replace '^HKLM:', 'HKEY_LOCAL_MACHINE\'
            $key = $key -replace '^HKCR:', 'HKEY_CLASSES_ROOT\'
            $key = $key -replace '^HKCU:', 'HKEY_CURRENT_USER\'
            $key = $key -replace '^HKU:', 'HKEY_USERS\'
            $key = $key -replace '^HKCC:', 'HKEY_CURRENT_CONFIG\'
            $key = $key -replace '^HKPD:', 'HKEY_PERFORMANCE_DATA\'
        }
        ElseIf ($Key -match '^HKLM\\|^HKCU\\|^HKCR\\|^HKU\\|^HKCC\\|^HKPD\\')
        {
            #  Converts registry paths that start with, e.g.: HKLM\
            $key = $key -replace '^HKLM\\', 'HKEY_LOCAL_MACHINE\'
            $key = $key -replace '^HKCR\\', 'HKEY_CLASSES_ROOT\'
            $key = $key -replace '^HKCU\\', 'HKEY_CURRENT_USER\'
            $key = $key -replace '^HKU\\', 'HKEY_USERS\'
            $key = $key -replace '^HKCC\\', 'HKEY_CURRENT_CONFIG\'
            $key = $key -replace '^HKPD\\', 'HKEY_PERFORMANCE_DATA\'
        }
            
        If ($PSBoundParameters.ContainsKey('SID'))
        {
            ## If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID				
            If ($key -match '^HKEY_CURRENT_USER\\') { $key = $key -replace '^HKEY_CURRENT_USER\\', "HKEY_USERS\$SID\" }
        }
            
        ## Append the PowerShell drive to the registry key path
        If ($key -notmatch '^Registry::') { [string]$key = "Registry::$key" }
            
        If ($Key -match '^Registry::HKEY_LOCAL_MACHINE|^Registry::HKEY_CLASSES_ROOT|^Registry::HKEY_CURRENT_USER|^Registry::HKEY_USERS|^Registry::HKEY_CURRENT_CONFIG|^Registry::HKEY_PERFORMANCE_DATA')
        {
            ## Check for expected key string format
            Write-Log -Message "Return fully qualified registry key path [$key]." -Source ${CmdletName}
            Write-Output -InputObject $key
        }
        Else
        {
            #  If key string is not properly formatted, throw an error
            Throw "Unable to detect target registry hive in string [$key]."
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Test-RegistryValue (From PowerShell App Deploy Toolkit)
Function Test-RegistryValue
{
    <#
    .SYNOPSIS
        Test if a registry value exists.
    .DESCRIPTION
        Checks a registry key path to see if it has a value with a given name. Can correctly handle cases where a value simply has an empty or null value.
    .PARAMETER Key
        Path of the registry key.
    .PARAMETER Value
        Specify the registry key value to check the existence of.
    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.
        Specify this parameter from the Invoke-HKCURegistrySettingsForAllUsers function to read/edit HKCU registry settings for all users on the system.
    .EXAMPLE
        Test-RegistryValue -Key 'HKLM:SYSTEM\CurrentControlSet\Control\Session Manager' -Value 'PendingFileRenameOperations'
    .NOTES
        To test if registry key exists, use Test-Path function like so:
        Test-Path -Path $Key -PathType 'Container'
    .LINK
        http://psappdeploytoolkit.com
    #>
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]$Key,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]$Value,
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullorEmpty()]
        [string]$SID
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        ## If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID
        Try
        {
            If ($PSBoundParameters.ContainsKey('SID'))
            {
                [string]$Key = Convert-RegistryPath -Key $Key -SID $SID
            }
            Else
            {
                [string]$Key = Convert-RegistryPath -Key $Key
            }
        }
        Catch
        {
            Throw
        }
        [boolean]$IsRegistryValueExists = $false
        Try
        {
            If (Test-Path -LiteralPath $Key -ErrorAction 'Stop')
            {
                [string[]]$PathProperties = Get-Item -LiteralPath $Key -ErrorAction 'Stop' | Select-Object -ExpandProperty 'Property' -ErrorAction 'Stop'
                If ($PathProperties -contains $Value) { $IsRegistryValueExists = $true }
            }
        }
        Catch { }
            
        If ($IsRegistryValueExists)
        {
            Write-Log -Message "Registry key value [$Key] [$Value] does exist." -Source ${CmdletName}
        }
        Else
        {
            Write-Log -Message "Registry key value [$Key] [$Value] does not exist." -Source ${CmdletName}
        }
        Write-Output -InputObject $IsRegistryValueExists
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Get-RegistryKey (From PowerShell App Deploy Toolkit)
Function Get-RegistryKey
{
    <#
    .SYNOPSIS
        Retrieves value names and value data for a specified registry key or optionally, a specific value.
    .DESCRIPTION
        Retrieves value names and value data for a specified registry key or optionally, a specific value.
        If the registry key does not exist or contain any values, the function will return $null by default. To test for existence of a registry key path, use built-in Test-Path cmdlet.
    .PARAMETER Key
        Path of the registry key.
    .PARAMETER Value
        Value to retrieve (optional).
    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.
        Specify this parameter from the Invoke-HKCURegistrySettingsForAllUsers function to read/edit HKCU registry settings for all users on the system.
    .PARAMETER ReturnEmptyKeyIfExists
        Return the registry key if it exists but it has no property/value pairs underneath it. Default is: $false.
    .PARAMETER DoNotExpandEnvironmentNames
        Return unexpanded REG_EXPAND_SZ values. Default is: $false.	
    .PARAMETER ContinueOnError
        Continue if an error is encountered. Default is: $true.
    .EXAMPLE
        Get-RegistryKey -Key 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1AD147D0-BE0E-3D6C-AC11-64F6DC4163F1}'
    .EXAMPLE
        Get-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\iexplore.exe'
    .EXAMPLE
        Get-RegistryKey -Key 'HKLM:Software\Wow6432Node\Microsoft\Microsoft SQL Server Compact Edition\v3.5' -Value 'Version'
    .EXAMPLE
        Get-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -Value 'Path' -DoNotExpandEnvironmentNames 
        Returns %ProgramFiles%\Java instead of C:\Program Files\Java
    .EXAMPLE
        Get-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Example' -Value '(Default)'
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Key,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [string]$SID,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [switch]$ReturnEmptyKeyIfExists = $false,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [switch]$DoNotExpandEnvironmentNames = $false,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [boolean]$ContinueOnError = $true
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            ## If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID
            If ($PSBoundParameters.ContainsKey('SID'))
            {
                [string]$key = Convert-RegistryPath -Key $key -SID $SID
            }
            Else
            {
                [string]$key = Convert-RegistryPath -Key $key
            }
                
            ## Check if the registry key exists
            If (-not (Test-Path -LiteralPath $key -ErrorAction 'Stop'))
            {
                Write-Log -Message "Registry key [$key] does not exist. Return `$null." -Severity 2 -Source ${CmdletName}
                $regKeyValue = $null
            }
            Else
            {
                If ($PSBoundParameters.ContainsKey('Value'))
                {
                    Write-Log -Message "Get registry key [$key] value [$value]." -Source ${CmdletName}
                }
                Else
                {
                    Write-Log -Message "Get registry key [$key] and all property values." -Source ${CmdletName}
                }
                    
                ## Get all property values for registry key
                $regKeyValue = Get-ItemProperty -LiteralPath $key -ErrorAction 'Stop'
                [int32]$regKeyValuePropertyCount = $regKeyValue | Measure-Object | Select-Object -ExpandProperty 'Count'
                    
                ## Select requested property
                If ($PSBoundParameters.ContainsKey('Value'))
                {
                    #  Check if registry value exists
                    [boolean]$IsRegistryValueExists = $false
                    If ($regKeyValuePropertyCount -gt 0)
                    {
                        Try
                        {
                            [string[]]$PathProperties = Get-Item -LiteralPath $Key -ErrorAction 'Stop' | Select-Object -ExpandProperty 'Property' -ErrorAction 'Stop'
                            If ($PathProperties -contains $Value) { $IsRegistryValueExists = $true }
                        }
                        Catch { }
                    }
                        
                    #  Get the Value (do not make a strongly typed variable because it depends entirely on what kind of value is being read)
                    If ($IsRegistryValueExists)
                    {
                        If ($DoNotExpandEnvironmentNames)
                        {
                            #Only useful on 'ExpandString' values
                            If ($Value -like '(Default)')
                            {
                                $regKeyValue = $(Get-Item -LiteralPath $key -ErrorAction 'Stop').GetValue($null, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                            }
                            Else
                            {
                                $regKeyValue = $(Get-Item -LiteralPath $key -ErrorAction 'Stop').GetValue($Value, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)	
                            }							
                        }
                        ElseIf ($Value -like '(Default)')
                        {
                            $regKeyValue = $(Get-Item -LiteralPath $key -ErrorAction 'Stop').GetValue($null)
                        }
                        Else
                        {
                            $regKeyValue = $regKeyValue | Select-Object -ExpandProperty $Value -ErrorAction 'SilentlyContinue'
                        }
                    }
                    Else
                    {
                        Write-Log -Message "Registry key value [$Key] [$Value] does not exist. Return `$null." -Source ${CmdletName}
                        $regKeyValue = $null
                    }
                }
                ## Select all properties or return empty key object
                Else
                {
                    If ($regKeyValuePropertyCount -eq 0)
                    {
                        If ($ReturnEmptyKeyIfExists)
                        {
                            Write-Log -Message "No property values found for registry key. Return empty registry key object [$key]." -Source ${CmdletName}
                            $regKeyValue = Get-Item -LiteralPath $key -Force -ErrorAction 'Stop'
                        }
                        Else
                        {
                            Write-Log -Message "No property values found for registry key. Return `$null." -Source ${CmdletName}
                            $regKeyValue = $null
                        }
                    }
                }
            }
            Write-Output -InputObject ($regKeyValue)
        }
        Catch
        {
            If (-not $Value)
            {
                Write-Log -Message "Failed to read registry key [$key]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
                If (-not $ContinueOnError)
                {
                    Throw "Failed to read registry key [$key]: $($_.Exception.Message)"
                }
            }
            Else
            {
                Write-Log -Message "Failed to read registry key [$key] value [$value]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
                If (-not $ContinueOnError)
                {
                    Throw "Failed to read registry key [$key] value [$value]: $($_.Exception.Message)"
                }
            }
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Set-RegistryKey (From PowerShell App Deploy Toolkit)
Function Set-RegistryKey
{
    <#
    .SYNOPSIS
        Creates a registry key name, value, and value data; it sets the same if it already exists.
    .DESCRIPTION
        Creates a registry key name, value, and value data; it sets the same if it already exists.
    .PARAMETER Key
        The registry key path.
    .PARAMETER Name
        The value name.
    .PARAMETER Value
        The value data.
    .PARAMETER Type
        The type of registry value to create or set. Options: 'Binary','DWord','ExpandString','MultiString','None','QWord','String','Unknown'. Default: String.
        Dword should be specified as a decimal.
    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.
        Specify this parameter from the Invoke-HKCURegistrySettingsForAllUsers function to read/edit HKCU registry settings for all users on the system.
    .PARAMETER ContinueOnError
        Continue if an error is encountered. Default is: $true.
    .EXAMPLE
        Set-RegistryKey -Key $blockedAppPath -Name 'Debugger' -Value $blockedAppDebuggerValue
    .EXAMPLE
        Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE' -Name 'Application' -Type 'Dword' -Value '1'
    .EXAMPLE
        Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name 'Debugger' -Value $blockedAppDebuggerValue -Type String
    .EXAMPLE
        Set-RegistryKey -Key 'HKCU\Software\Microsoft\Example' -Name 'Data' -Value (0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x02,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x02,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x00,0x01,0x01,0x01,0x02,0x02,0x02) -Type 'Binary'
    .EXAMPLE
        Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Example' -Value '(Default)'
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Key,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        $Value,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Binary', 'DWord', 'ExpandString', 'MultiString', 'None', 'QWord', 'String', 'Unknown')]
        [Microsoft.Win32.RegistryValueKind]$Type = 'String',
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [string]$SID,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [boolean]$ContinueOnError = $true
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            [string]$RegistryValueWriteAction = 'set'
                
            ## If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID
            If ($PSBoundParameters.ContainsKey('SID'))
            {
                [string]$key = Convert-RegistryPath -Key $key -SID $SID
            }
            Else
            {
                [string]$key = Convert-RegistryPath -Key $key
            }
                
            ## Create registry key if it doesn't exist
            If (-not (Test-Path -LiteralPath $key -ErrorAction 'Stop'))
            {
                Try
                {
                    Write-Log -Message "Create registry key [$key]." -Source ${CmdletName}
                    # No forward slash found in Key. Use New-Item cmdlet to create registry key
                    If ((($Key -split '/').Count - 1) -eq 0)
                    {
                        $null = New-Item -Path $key -ItemType 'Registry' -Force -ErrorAction 'Stop'
                    }
                    # Forward slash was found in Key. Use REG.exe ADD to create registry key 
                    Else
                    {
                        [string]$CreateRegkeyResult = & reg.exe Add "$($Key.Substring($Key.IndexOf('::') + 2))"
                        If ($global:LastExitCode -ne 0)
                        {
                            Throw "Failed to create registry key [$Key]"
                        }
                    }
                }
                Catch
                {
                    Throw
                }
            }
                
            If ($Name)
            {
                ## Set registry value if it doesn't exist
                If (-not (Get-ItemProperty -LiteralPath $key -Name $Name -ErrorAction 'SilentlyContinue'))
                {
                    Write-Log -Message "Set registry key value: [$key] [$name = $value]." -Source ${CmdletName}
                    $null = New-ItemProperty -LiteralPath $key -Name $name -Value $value -PropertyType $Type -ErrorAction 'Stop'
                }
                ## Update registry value if it does exist
                Else
                {
                    [string]$RegistryValueWriteAction = 'update'
                    If ($Name -eq '(Default)')
                    {
                        ## Set Default registry key value with the following workaround, because Set-ItemProperty contains a bug and cannot set Default registry key value
                        $null = $(Get-Item -LiteralPath $key -ErrorAction 'Stop').OpenSubKey('', 'ReadWriteSubTree').SetValue($null, $value)
                    } 
                    Else
                    {
                        Write-Log -Message "Update registry key value: [$key] [$name = $value]." -Source ${CmdletName}
                        $null = Set-ItemProperty -LiteralPath $key -Name $name -Value $value -ErrorAction 'Stop'
                    }
                }
            }
        }
        Catch
        {
            If ($Name)
            {
                Write-Log -Message "Failed to $RegistryValueWriteAction value [$value] for registry key [$key] [$name]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
                If (-not $ContinueOnError)
                {
                    Throw "Failed to $RegistryValueWriteAction value [$value] for registry key [$key] [$name]: $($_.Exception.Message)"
                }
            }
            Else
            {
                Write-Log -Message "Failed to set registry key [$key]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
                If (-not $ContinueOnError)
                {
                    Throw "Failed to set registry key [$key]: $($_.Exception.Message)"
                }
            }
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Remove-RegistryKey (From PowerShell App Deploy Toolkit)
Function Remove-RegistryKey
{
    <#
    .SYNOPSIS
        Deletes the specified registry key or value.
    .DESCRIPTION
        Deletes the specified registry key or value.
    .PARAMETER Key
        Path of the registry key to delete.
    .PARAMETER Name
        Name of the registry value to delete.
    .PARAMETER Recurse
        Delete registry key recursively.
    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.
        Specify this parameter from the Invoke-HKCURegistrySettingsForAllUsers function to read/edit HKCU registry settings for all users on the system.
    .PARAMETER ContinueOnError
        Continue if an error is encountered. Default is: $true.
    .EXAMPLE
        Remove-RegistryKey -Key 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce'
    .EXAMPLE
        Remove-RegistryKey -Key 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'RunAppInstall'
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Key,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [switch]$Recurse,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [string]$SID,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [boolean]$ContinueOnError = $true
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            ## If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID
            If ($PSBoundParameters.ContainsKey('SID'))
            {
                [string]$Key = Convert-RegistryPath -Key $Key -SID $SID
            }
            Else
            {
                [string]$Key = Convert-RegistryPath -Key $Key
            }
                
            If (-not ($Name))
            {
                If (Test-Path -LiteralPath $Key -ErrorAction 'Stop')
                {
                    If ($Recurse)
                    {
                        Write-Log -Message "Delete registry key recursively [$Key]." -Source ${CmdletName}
                        $null = Remove-Item -LiteralPath $Key -Force -Recurse -ErrorAction 'Stop'
                    }
                    Else
                    {
                        If ($null -eq (Get-ChildItem -LiteralPath $Key -ErrorAction 'Stop'))
                        {
                            ## Check if there are subkeys of $Key, if so, executing Remove-Item will hang. Avoiding this with Get-ChildItem.
                            Write-Log -Message "Delete registry key [$Key]." -Source ${CmdletName}
                            $null = Remove-Item -LiteralPath $Key -Force -ErrorAction 'Stop'
                        }
                        Else
                        {
                            Throw "Unable to delete child key(s) of [$Key] without [-Recurse] switch."
                        }
                    }
                }
                Else
                {
                    Write-Log -Message "Unable to delete registry key [$Key] because it does not exist." -Severity 2 -Source ${CmdletName}
                }
            }
            Else
            {
                If (Test-Path -LiteralPath $Key -ErrorAction 'Stop')
                {
                    Write-Log -Message "Delete registry value [$Key] [$Name]." -Source ${CmdletName}
                        
                    If ($Name -eq '(Default)')
                    {
                        ## Remove (Default) registry key value with the following workaround because Remove-ItemProperty cannot remove the (Default) registry key value
                        $null = (Get-Item -LiteralPath $Key -ErrorAction 'Stop').OpenSubKey('', 'ReadWriteSubTree').DeleteValue('')
                    }
                    Else
                    {
                        $null = Remove-ItemProperty -LiteralPath $Key -Name $Name -Force -ErrorAction 'Stop'
                    }
                }
                Else
                {
                    Write-Log -Message "Unable to delete registry value [$Key] [$Name] because registry key does not exist." -Severity 2 -Source ${CmdletName}
                }
            }
        }
        Catch [System.Management.Automation.PSArgumentException]
        {
            Write-Log -Message "Unable to delete registry value [$Key] [$Name] because it does not exist." -Severity 2 -Source ${CmdletName}
        }
        Catch
        {
            If (-not ($Name))
            {
                Write-Log -Message "Failed to delete registry key [$Key]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
                If (-not $ContinueOnError)
                {
                    Throw "Failed to delete registry key [$Key]: $($_.Exception.Message)"
                }
            }
            Else
            {
                Write-Log -Message "Failed to delete registry value [$Key] [$Name]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
                If (-not $ContinueOnError)
                {
                    Throw "Failed to delete registry value [$Key] [$Name]: $($_.Exception.Message)"
                }
            }
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Invoke-HKCURegistrySettingsForAllUsers (From PowerShell App Deploy Toolkit)
Function Invoke-HKCURegistrySettingsForAllUsers
{
    <#
    .SYNOPSIS
        Set current user registry settings for all current users and any new users in the future.
    .DESCRIPTION
        Set HKCU registry settings for all current and future users by loading their NTUSER.dat registry hive file, and making the modifications.
        This function will modify HKCU settings for all users even when executed under the SYSTEM account.
        To ensure new users in the future get the registry edits, the Default User registry hive used to provision the registry for new users is modified.
        This function can be used as an alternative to using ActiveSetup for registry settings.
        The advantage of using this function over ActiveSetup is that a user does not have to log off and log back on before the changes take effect.
    .PARAMETER RegistrySettings
        Script block which contains HKCU registry settings which should be modified for all users on the system. Must specify the -SID parameter for all HKCU settings.
    .PARAMETER UserProfiles
        Specify the user profiles to modify HKCU registry settings for. Default is all user profiles except for system profiles.
    .EXAMPLE
        [scriptblock]$HKCURegistrySettings = {
            Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'qmenable' -Value 0 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'updatereliabilitydata' -Value 1 -Type DWord -SID $UserProfile.SID
        }
        Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [scriptblock]$RegistrySettings,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [psobject[]]$UserProfiles = (Get-UserProfiles)
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        ForEach ($UserProfile in $UserProfiles)
        {
            Try
            {
                #  Set the path to the user's registry hive when it is loaded
                [string]$UserRegistryPath = "Registry::HKEY_USERS\$($UserProfile.SID)"
                    
                #  Set the path to the user's registry hive file
                [string]$UserRegistryHiveFile = Join-Path -Path $UserProfile.ProfilePath -ChildPath 'NTUSER.DAT'
                    
                #  Load the User profile registry hive if it is not already loaded because the User is logged in
                [boolean]$ManuallyLoadedRegHive = $false
                If (-not (Test-Path -LiteralPath $UserRegistryPath))
                {
                    #  Load the User registry hive if the registry hive file exists
                    If (Test-Path -LiteralPath $UserRegistryHiveFile -PathType 'Leaf')
                    {
                        Write-Log -Message "Load the User [$($UserProfile.NTAccount)] registry hive in path [HKEY_USERS\$($UserProfile.SID)]." -Source ${CmdletName}
                        [string]$HiveLoadResult = & reg.exe load "`"HKEY_USERS\$($UserProfile.SID)`"" "`"$UserRegistryHiveFile`""
                            
                        If ($global:LastExitCode -ne 0)
                        {
                            Throw "Failed to load the registry hive for User [$($UserProfile.NTAccount)] with SID [$($UserProfile.SID)]. Failure message [$HiveLoadResult]. Continue..."
                        }
                            
                        [boolean]$ManuallyLoadedRegHive = $true
                    }
                    Else
                    {
                        Throw "Failed to find the registry hive file [$UserRegistryHiveFile] for User [$($UserProfile.NTAccount)] with SID [$($UserProfile.SID)]. Continue..."
                    }
                }
                Else
                {
                    Write-Log -Message "The User [$($UserProfile.NTAccount)] registry hive is already loaded in path [HKEY_USERS\$($UserProfile.SID)]." -Source ${CmdletName}
                }
                    
                ## Execute ScriptBlock which contains code to manipulate HKCU registry.
                #  Make sure read/write calls to the HKCU registry hive specify the -SID parameter or settings will not be changed for all users.
                #  Example: Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'qmenable' -Value 0 -Type DWord -SID $UserProfile.SID
                Write-Log -Message 'Execute ScriptBlock to modify HKCU registry settings for all users.' -Source ${CmdletName}
                & $RegistrySettings
            }
            Catch
            {
                Write-Log -Message "Failed to modify the registry hive for User [$($UserProfile.NTAccount)] with SID [$($UserProfile.SID)] `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            }
            Finally
            {
                If ($ManuallyLoadedRegHive)
                {
                    Try
                    {
                        Write-Log -Message "Unload the User [$($UserProfile.NTAccount)] registry hive in path [HKEY_USERS\$($UserProfile.SID)]." -Source ${CmdletName}
                        [string]$HiveLoadResult = & reg.exe unload "`"HKEY_USERS\$($UserProfile.SID)`""
                            
                        If ($global:LastExitCode -ne 0)
                        {
                            Write-Log -Message "REG.exe failed to unload the registry hive and exited with exit code [$($global:LastExitCode)]. Performing manual garbage collection to ensure successful unloading of registry hive." -Severity 2 -Source ${CmdletName}
                            [GC]::Collect()
                            [GC]::WaitForPendingFinalizers()
                            Start-Sleep -Seconds 5
                                
                            Write-Log -Message "Unload the User [$($UserProfile.NTAccount)] registry hive in path [HKEY_USERS\$($UserProfile.SID)]." -Source ${CmdletName}
                            [string]$HiveLoadResult = & reg.exe unload "`"HKEY_USERS\$($UserProfile.SID)`""
                            If ($global:LastExitCode -ne 0) { Throw "REG.exe failed with exit code [$($global:LastExitCode)] and result [$HiveLoadResult]." }
                        }
                    }
                    Catch
                    {
                        Write-Log -Message "Failed to unload the registry hive for User [$($UserProfile.NTAccount)] with SID [$($UserProfile.SID)]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
                    }
                }
            }
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function ConvertTo-NTAccountOrSID (From PowerShell App Deploy Toolkit)
Function ConvertTo-NTAccountOrSID
{
    <#
    .SYNOPSIS
        Convert between NT Account names and their security identifiers (SIDs).
    .DESCRIPTION
        Specify either the NT Account name or the SID and get the other. Can also convert well known sid types.
    .PARAMETER AccountName
        The Windows NT Account name specified in <domain>\<username> format.
        Use fully qualified account names (e.g., <domain>\<username>) instead of isolated names (e.g, <username>) because they are unambiguous and provide better performance.
    .PARAMETER SID
        The Windows NT Account SID.
    .PARAMETER WellKnownSIDName
        Specify the Well Known SID name translate to the actual SID (e.g., LocalServiceSid).
        To get all well known SIDs available on system: [enum]::GetNames([Security.Principal.WellKnownSidType])
    .PARAMETER WellKnownToNTAccount
        Convert the Well Known SID to an NTAccount name
    .EXAMPLE
        ConvertTo-NTAccountOrSID -AccountName 'CONTOSO\User1'
        Converts a Windows NT Account name to the corresponding SID
    .EXAMPLE
        ConvertTo-NTAccountOrSID -SID 'S-1-5-21-1220945662-2111687655-725345543-14012660'
        Converts a Windows NT Account SID to the corresponding NT Account Name
    .EXAMPLE
        ConvertTo-NTAccountOrSID -WellKnownSIDName 'NetworkServiceSid'
        Converts a Well Known SID name to a SID
    .NOTES
        This is an internal script function and should typically not be called directly.
        The conversion can return an empty result if the user account does not exist anymore or if translation fails.
        http://blogs.technet.com/b/askds/archive/2011/07/28/troubleshooting-sid-translation-failures-from-the-obvious-to-the-not-so-obvious.aspx
    .LINK
        http://psappdeploytoolkit.com
        List of Well Known SIDs: http://msdn.microsoft.com/en-us/library/system.security.principal.wellknownsidtype(v=vs.110).aspx
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ParameterSetName = 'NTAccountToSID', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$AccountName,
        [Parameter(Mandatory = $true, ParameterSetName = 'SIDToNTAccount', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SID,
        [Parameter(Mandatory = $true, ParameterSetName = 'WellKnownName', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WellKnownSIDName,
        [Parameter(Mandatory = $false, ParameterSetName = 'WellKnownName')]
        [ValidateNotNullOrEmpty()]
        [switch]$WellKnownToNTAccount
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            Switch ($PSCmdlet.ParameterSetName)
            {
                'SIDToNTAccount'
                {
                    [string]$msg = "the SID [$SID] to an NT Account name"
                    Write-Log -Message "Convert $msg." -Source ${CmdletName}
                        
                    $NTAccountSID = New-Object -TypeName 'System.Security.Principal.SecurityIdentifier' -ArgumentList $SID
                    $NTAccount = $NTAccountSID.Translate([Security.Principal.NTAccount])
                    Write-Output -InputObject $NTAccount
                }
                'NTAccountToSID'
                {
                    [string]$msg = "the NT Account [$AccountName] to a SID"
                    Write-Log -Message "Convert $msg." -Source ${CmdletName}
                        
                    $NTAccount = New-Object -TypeName 'System.Security.Principal.NTAccount' -ArgumentList $AccountName
                    $NTAccountSID = $NTAccount.Translate([Security.Principal.SecurityIdentifier])
                    Write-Output -InputObject $NTAccountSID
                }
                'WellKnownName'
                {
                    If ($WellKnownToNTAccount)
                    {
                        [string]$ConversionType = 'NTAccount'
                    }
                    Else
                    {
                        [string]$ConversionType = 'SID'
                    }
                    [string]$msg = "the Well Known SID Name [$WellKnownSIDName] to a $ConversionType"
                    Write-Log -Message "Convert $msg." -Source ${CmdletName}
                        
                    #  Get the SID for the root domain
                    Try
                    {
                        $MachineRootDomain = (Get-WmiObject -Class 'Win32_ComputerSystem' -ErrorAction 'Stop').Domain.ToLower()
                        $ADDomainObj = New-Object -TypeName 'System.DirectoryServices.DirectoryEntry' -ArgumentList "LDAP://$MachineRootDomain"
                        $DomainSidInBinary = $ADDomainObj.ObjectSid
                        $DomainSid = New-Object -TypeName 'System.Security.Principal.SecurityIdentifier' -ArgumentList ($DomainSidInBinary[0], 0)
                    }
                    Catch
                    {
                        Write-Log -Message 'Unable to get Domain SID from Active Directory. Setting Domain SID to $null.' -Severity 2 -Source ${CmdletName}
                        $DomainSid = $null
                    }
                        
                    #  Get the SID for the well known SID name
                    $WellKnownSidType = [Security.Principal.WellKnownSidType]::$WellKnownSIDName
                    $NTAccountSID = New-Object -TypeName 'System.Security.Principal.SecurityIdentifier' -ArgumentList ($WellKnownSidType, $DomainSid)
                        
                    If ($WellKnownToNTAccount)
                    {
                        $NTAccount = $NTAccountSID.Translate([Security.Principal.NTAccount])
                        Write-Output -InputObject $NTAccount
                    }
                    Else
                    {
                        Write-Output -InputObject $NTAccountSID
                    }
                }
            }
        }
        Catch
        {
            Write-Log -Message "Failed to convert $msg. It may not be a valid account anymore or there is some other problem. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Get-UserProfiles (From PowerShell App Deploy Toolkit)
Function Get-UserProfiles
{
    <#
    .SYNOPSIS
        Get the User Profile Path, User Account Sid, and the User Account Name for all users that log onto the machine and also the Default User (which does not log on).
    .DESCRIPTION
        Get the User Profile Path, User Account Sid, and the User Account Name for all users that log onto the machine and also the Default User (which does  not log on).
        Please note that the NTAccount property may be empty for some user profiles but the SID and ProfilePath properties will always be populated.
    .PARAMETER ExcludeNTAccount
        Specify NT account names in Domain\Username format to exclude from the list of user profiles.
    .PARAMETER ExcludeSystemProfiles
        Exclude system profiles: SYSTEM, LOCAL SERVICE, NETWORK SERVICE. Default is: $true.
    .PARAMETER ExcludeDefaultUser
        Exclude the Default User. Default is: $false.
    .EXAMPLE
        Get-UserProfiles
        Returns the following properties for each user profile on the system: NTAccount, SID, ProfilePath
    .EXAMPLE
        Get-UserProfiles -ExcludeNTAccount 'CONTOSO\Robot','CONTOSO\ntadmin'
    .EXAMPLE
        [string[]]$ProfilePaths = Get-UserProfiles | Select-Object -ExpandProperty 'ProfilePath'
        Returns the user profile path for each user on the system. This information can then be used to make modifications under the user profile on the filesystem.
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ExcludeNTAccount,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [boolean]$ExcludeSystemProfiles = $true,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [switch]$ExcludeDefaultUser = $false
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            Write-Log -Message 'Get the User Profile Path, User Account SID, and the User Account Name for all users that log onto the machine.' -Source ${CmdletName}
                
            ## Get the User Profile Path, User Account Sid, and the User Account Name for all users that log onto the machine
            [string]$UserProfileListRegKey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
            [psobject[]]$UserProfiles = Get-ChildItem -LiteralPath $UserProfileListRegKey -ErrorAction 'Stop' |
            ForEach-Object {
                Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction 'Stop' | Where-Object { ($_.ProfileImagePath) } |
                Select-Object @{ Label = 'NTAccount'; Expression = { $(ConvertTo-NTAccountOrSID -SID $_.PSChildName).Value } }, @{ Label = 'SID'; Expression = { $_.PSChildName } }, @{ Label = 'ProfilePath'; Expression = { $_.ProfileImagePath } }
            }
            If ($ExcludeSystemProfiles)
            {
                [string[]]$SystemProfiles = 'S-1-5-18', 'S-1-5-19', 'S-1-5-20'
                [psobject[]]$UserProfiles = $UserProfiles | Where-Object { $SystemProfiles -notcontains $_.SID }
            }
            If ($ExcludeNTAccount)
            {
                [psobject[]]$UserProfiles = $UserProfiles | Where-Object { $ExcludeNTAccount -notcontains $_.NTAccount }
            }
                
            ## Find the path to the Default User profile
            If (-not $ExcludeDefaultUser)
            {
                [string]$UserProfilesDirectory = Get-ItemProperty -LiteralPath $UserProfileListRegKey -Name 'ProfilesDirectory' -ErrorAction 'Stop' | Select-Object -ExpandProperty 'ProfilesDirectory'
                    
                #  On Windows Vista or higher
                If (([version]$envOSVersion).Major -gt 5)
                {
                    # Path to Default User Profile directory on Windows Vista or higher: By default, C:\Users\Default
                    [string]$DefaultUserProfileDirectory = Get-ItemProperty -LiteralPath $UserProfileListRegKey -Name 'Default' -ErrorAction 'Stop' | Select-Object -ExpandProperty 'Default'
                }
                #  On Windows XP or lower
                Else
                {
                    #  Default User Profile Name: By default, 'Default User'
                    [string]$DefaultUserProfileName = Get-ItemProperty -LiteralPath $UserProfileListRegKey -Name 'DefaultUserProfile' -ErrorAction 'Stop' | Select-Object -ExpandProperty 'DefaultUserProfile'
                        
                    #  Path to Default User Profile directory: By default, C:\Documents and Settings\Default User
                    [string]$DefaultUserProfileDirectory = Join-Path -Path $UserProfilesDirectory -ChildPath $DefaultUserProfileName
                }
                    
                ## Create a custom object for the Default User profile.
                #  Since the Default User is not an actual User account, it does not have a username or a SID.
                #  We will make up a SID and add it to the custom object so that we have a location to load the default registry hive into later on.
                [psobject]$DefaultUserProfile = New-Object -TypeName 'PSObject' -Property @{
                    NTAccount   = 'Default User'
                    SID         = 'S-1-5-21-Default-User'
                    ProfilePath = $DefaultUserProfileDirectory
                }
                    
                ## Add the Default User custom object to the User Profile list.
                $UserProfiles += $DefaultUserProfile
            }
                
            Write-Output -InputObject $UserProfiles
        }
        Catch
        {
            Write-Log -Message "Failed to create a custom object representing all user profiles on the machine. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion

#region Function Update-Desktop (From PowerShell App Deploy Toolkit)
Function Update-Desktop
{
    <#
    .SYNOPSIS
        Refresh the Windows Explorer Shell, which causes the desktop icons and the environment variables to be reloaded.
    .DESCRIPTION
        Refresh the Windows Explorer Shell, which causes the desktop icons and the environment variables to be reloaded.
    .PARAMETER ContinueOnError
        Continue if an error is encountered. Default is: $true.
    .EXAMPLE
        Update-Desktop
    .NOTES
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [boolean]$ContinueOnError = $true
    )
        
    Begin
    {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process
    {
        Try
        {
            Write-Log -Message 'Refresh the Desktop and the Windows Explorer environment process block.' -Source ${CmdletName}
            [PSADT.Explorer]::RefreshDesktopAndEnvironmentVariables()
        }
        Catch
        {
            Write-Log -Message "Failed to refresh the Desktop and the Windows Explorer environment process block. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            If (-not $ContinueOnError)
            {
                Throw "Failed to refresh the Desktop and the Windows Explorer environment process block: $($_.Exception.Message)"
            }
        }
    }
    End
    {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
Set-Alias -Name 'Refresh-Desktop' -Value 'Update-Desktop' -Scope 'Script' -Force -ErrorAction 'SilentlyContinue'
#endregion