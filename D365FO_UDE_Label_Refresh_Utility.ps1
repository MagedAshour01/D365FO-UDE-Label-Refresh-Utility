#requires -Version 5.1
<#
.SYNOPSIS
    Compile RGSLAI Core and UX labels into a NEW staging folder, not the project.
.DESCRIPTION
    Uses the configured-reference bin\LabelC.exe by default. A different compiler
    must be explicitly supplied with -CompilerPath; no compiler from another
    version is silently selected. Records compiler identity, exit codes, stdout,
    stderr, native logs, source hashes, and output hashes in an evidence ZIP.

    Does NOT install outputs, change source XML/TXT, edit configurations, perform
    full builds, deploy, synchronize databases, restart services, or change policy.
    No administrator privilege is requested. Permissions/policy failures stop work.

    RESULT=STAGED_DLLS_REQUIRE_REVIEW means nonempty DLLs were generated with exit
    code zero. It does NOT prove embedded label correctness or runtime resolution.

    Designed for Windows PowerShell 5.1. The actual D365 compiler and Windows
    execution could not be tested in the authoring environment.
.EXAMPLE
    & 'D:\Temp\RGSLAI_Label_Compile_Stage.ps1'
.EXAMPLE
    & 'D:\Temp\RGSLAI_Label_Compile_Stage.ps1' -CompilerPath 'D:\VerifiedTools\LabelC.exe'
    Use the override only after verifying that it belongs to the active toolset.
#>
[CmdletBinding()]
param(
    [string]$MetadataRoot = 'D:\RGSLAI Project\RGSLAI',
    [string]$ReferencePackagesRoot = 'D:\Dynamics365\10.0.2428.63\PackagesLocalDirectory',
    [string]$CompilerPath,
    [string]$OutputParent = 'D:\Temp',
    [ValidateRange(1, 120)]
    [int]$CompilerTimeoutMinutes = 10
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$runRoot = $null
$result = 'NOT_STARTED'
$failureMessage = ''
$sourceBefore = @()
$moduleResults = @()
$compilerIdentity = $null
$writeOutputZip = $false

function Get-AbsoluteLocalPath {
    param([Parameter(Mandatory = $true)][string]$Value)
    if ($Value -notmatch '^[A-Za-z]:[\\/]' -or $Value -match '["\r\n]' -or $Value -match '[*?]') {
        throw "Use an absolute local drive path without wildcards or quotes: $Value"
    }
    return [IO.Path]::GetFullPath($Value).TrimEnd([char[]]@('\', '/'))
}

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
}

function Write-JsonFile {
    param([string]$Path, [object]$Value)
    $json = ConvertTo-Json -InputObject $Value -Depth 10
    [IO.File]::WriteAllText($Path, $json, $utf8)
}

function Invoke-LabelCompiler {
    param([string]$Module, [string]$OutputPath, [string]$LogFolder)

    $stdoutPath = Join-Path $LogFolder 'stdout.txt'
    $stderrPath = Join-Path $LogFolder 'stderr.txt'
    $outLog = Join-Path $LogFolder 'labelc.log'
    $errLog = Join-Path $LogFolder 'labelc.err.xml'
    # Use ProcessStartInfo rather than a shell or Invoke-Expression. No trailing
    # slash is included in quoted paths, avoiding Windows native quoting issues.
    $arguments = '-metadata="{0}" -modelmodule="{1}" -output="{2}" -outlog="{3}" -errlog="{4}"' -f $MetadataRoot, $Module, $OutputPath, $outLog, $errLog
    [IO.File]::WriteAllText((Join-Path $LogFolder 'invocation.txt'), ($CompilerPath + [Environment]::NewLine + $arguments), $utf8)

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $CompilerPath
    $startInfo.Arguments = $arguments
    $startInfo.WorkingDirectory = Split-Path -Parent $CompilerPath
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $timedOut = $false

    try {
        if (-not $process.Start()) { throw 'Label compiler could not be started.' }
        $stdOutTask = $process.StandardOutput.ReadToEndAsync()
        $stdErrTask = $process.StandardError.ReadToEndAsync()
        $deadline = [DateTime]::UtcNow.AddMinutes($CompilerTimeoutMinutes)
        while (-not $process.WaitForExit(1000)) {
            Write-Progress -Activity "Compile labels: $Module" -Status 'Writing to staging only; project source is unchanged.'
            if ([DateTime]::UtcNow -gt $deadline) {
                # Stop only the compiler process created by this invocation.
                $timedOut = $true
                $process.Kill()
                $process.WaitForExit()
                break
            }
        }
        $process.WaitForExit()
        [IO.File]::WriteAllText($stdoutPath, $stdOutTask.GetAwaiter().GetResult(), $utf8)
        [IO.File]::WriteAllText($stderrPath, $stdErrTask.GetAwaiter().GetResult(), $utf8)
        return [pscustomobject]@{ ExitCode = $process.ExitCode; TimedOut = $timedOut }
    }
    finally {
        Write-Progress -Activity "Compile labels: $Module" -Completed
        $process.Dispose()
    }
}

try {
    if ($env:OS -ne 'Windows_NT') { throw 'Run this script on the Windows development machine.' }
    $MetadataRoot = Get-AbsoluteLocalPath $MetadataRoot
    $ReferencePackagesRoot = Get-AbsoluteLocalPath $ReferencePackagesRoot
    $OutputParent = Get-AbsoluteLocalPath $OutputParent
    if (-not (Test-Path -LiteralPath $MetadataRoot -PathType Container)) { throw "Metadata root not found: $MetadataRoot" }
    if (-not (Test-Path -LiteralPath $ReferencePackagesRoot -PathType Container)) { throw "Reference root not found: $ReferencePackagesRoot" }
    if (-not (Test-Path -LiteralPath $OutputParent -PathType Container)) { throw "Output parent not found: $OutputParent" }
    foreach ($protectedRoot in @($MetadataRoot, $ReferencePackagesRoot)) {
        if ($OutputParent.Equals($protectedRoot, [StringComparison]::OrdinalIgnoreCase) -or
            $OutputParent.StartsWith(($protectedRoot + '\'), [StringComparison]::OrdinalIgnoreCase)) {
            throw 'OutputParent must be outside the source and reference package trees.'
        }
    }

    if ([string]::IsNullOrWhiteSpace($CompilerPath)) {
        $CompilerPath = Join-Path $ReferencePackagesRoot 'bin\LabelC.exe'
    }
    $CompilerPath = Get-AbsoluteLocalPath $CompilerPath
    if (-not (Test-Path -LiteralPath $CompilerPath -PathType Leaf)) {
        throw "LabelC.exe was not found at: $CompilerPath. No compiler was run. Supply -CompilerPath only after verifying the active toolset path; do not copy a compiler from another version."
    }
    if ([IO.Path]::GetFileName($CompilerPath) -ine 'LabelC.exe') { throw 'CompilerPath must point to LabelC.exe.' }
    $compilerFile = Get-Item -LiteralPath $CompilerPath
    $compilerIdentity = [pscustomobject]@{
        Path = $compilerFile.FullName
        FileVersion = $compilerFile.VersionInfo.FileVersion
        ProductVersion = $compilerFile.VersionInfo.ProductVersion
        SHA256 = Get-Sha256 $CompilerPath
        Note = 'Path is selected explicitly; file version is recorded, not assumed equal to application version.'
    }

    $modules = @('RGSLAIPlatformCore', 'RGSLAIPlatformUX')
    foreach ($module in $modules) {
        $packagePath = Join-Path $MetadataRoot $module
        $labelRoot = Join-Path $packagePath ($module + '\AxLabelFile')
        $descriptorPath = Join-Path $labelRoot ($module + '_en-US.xml')
        $resourcePath = Join-Path $labelRoot ('LabelResources\en-US\' + $module + '.en-US.label.txt')
        foreach ($requiredPath in @($descriptorPath, $resourcePath)) {
            if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) { throw "Required en-US source missing: $requiredPath" }
        }
        # Read source XML safely; do not rewrite or normalize it.
        $descriptor = New-Object System.Xml.XmlDocument
        $descriptor.XmlResolver = $null
        $descriptor.Load($descriptorPath)
        $idNode = $descriptor.SelectSingleNode('/AxLabelFile/LabelFileId')
        if ($null -eq $idNode -or $idNode.InnerText -cne $module) {
            throw "Unexpected LabelFileId in $descriptorPath. Source was not changed."
        }
        foreach ($sourceFile in @(Get-ChildItem -LiteralPath $labelRoot -File -Recurse)) {
            $sourceBefore += [pscustomobject]@{
                Module = $module; Path = $sourceFile.FullName
                SHA256 = Get-Sha256 $sourceFile.FullName
            }
        }
    }

    $runName = 'RGSLAI_LabelCompile_' + (Get-Date -Format 'yyyyMMdd_HHmmss_fff') + '_' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
    $runRoot = Join-Path $OutputParent $runName
    New-Item -Path $runRoot -ItemType Directory -ErrorAction Stop | Out-Null
    $writeOutputZip = $true
    Write-Output ('RGSLAI_LABEL_STAGE_DIRECTORY=' + $runRoot)
    Write-Output ('RGSLAI_LABEL_COMPILER=' + $CompilerPath)
    Write-JsonFile (Join-Path $runRoot 'compiler.json') $compilerIdentity
    Write-JsonFile (Join-Path $runRoot 'source-before.json') $sourceBefore

    foreach ($module in $modules) {
        $moduleStage = Join-Path $runRoot $module
        $outputPath = Join-Path $moduleStage 'Resources'
        $logsPath = Join-Path $moduleStage 'Logs'
        $evidencePath = Join-Path $moduleStage 'SourceEvidence'
        foreach ($newPath in @($moduleStage, $outputPath, $logsPath, $evidencePath)) {
            New-Item -Path $newPath -ItemType Directory -ErrorAction Stop | Out-Null
        }
        $sourceRoot = Join-Path $MetadataRoot ($module + '\' + $module + '\AxLabelFile')
        Copy-Item -LiteralPath (Join-Path $sourceRoot ($module + '_en-US.xml')) -Destination $evidencePath
        Copy-Item -LiteralPath (Join-Path $sourceRoot ('LabelResources\en-US\' + $module + '.en-US.label.txt')) -Destination $evidencePath

        Write-Output ('RGSLAI_LABEL_COMPILING=' + $module)
        $execution = Invoke-LabelCompiler -Module $module -OutputPath $outputPath -LogFolder $logsPath
        $dlls = @(Get-ChildItem -LiteralPath $outputPath -Recurse -File -Filter '*.dll' | Where-Object { $_.Length -gt 0 })
        $outputInventory = @(
            foreach ($file in @(Get-ChildItem -LiteralPath $outputPath -Recurse -File)) {
                [pscustomobject]@{
                    Path = $file.FullName; Bytes = $file.Length
                    LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
                    SHA256 = Get-Sha256 $file.FullName
                }
            }
        )
        Write-JsonFile (Join-Path $moduleStage 'output-inventory.json') $outputInventory
        $moduleResults += [pscustomobject]@{
            Module = $module; ExitCode = $execution.ExitCode; TimedOut = $execution.TimedOut
            NonEmptyDllCount = $dlls.Count; OutputPath = $outputPath
        }
        Write-Output ('RGSLAI_LABEL_EXIT_CODE_' + $module + '=' + $execution.ExitCode)
        Write-Output ('RGSLAI_LABEL_DLL_COUNT_' + $module + '=' + $dlls.Count)
        if ($execution.TimedOut) { throw "Label compiler timed out for $module. Review the staged logs; do not deploy." }
        if ($execution.ExitCode -ne 0) { throw "Label compiler failed for $module. Review the staged logs; do not deploy." }
        if ($dlls.Count -eq 0) { throw "Compiler returned zero but produced no nonempty DLLs for $module. Do not deploy." }
    }
    $result = 'STAGED_DLLS_REQUIRE_REVIEW'
}
catch {
    $result = 'FAILED_OR_BLOCKED'
    $failureMessage = $_.Exception.Message
    Write-Output ('RGSLAI_LABEL_ERROR=' + $failureMessage)
}
finally {
    $sourceUnchanged = $true
    foreach ($source in $sourceBefore) {
        try {
            if ((Get-Sha256 $source.Path) -cne $source.SHA256) { $sourceUnchanged = $false }
        }
        catch { $sourceUnchanged = $false }
    }
    if (-not $sourceUnchanged) {
        $result = 'SOURCE_CHANGED_OR_UNREADABLE_STOP'
        $failureMessage += ' Source hashes changed or could not be read. Do not install outputs or deploy.'
    }
    if ($writeOutputZip) {
        try {
            $summary = [pscustomobject]@{
                Result = $result; Error = $failureMessage; CheckUtc = [DateTime]::UtcNow.ToString('o')
                MetadataRoot = $MetadataRoot; ReferencePackagesRoot = $ReferencePackagesRoot
                Compiler = $compilerIdentity; Modules = $moduleResults
                ExistingLabelSourceFilesUnchanged = $sourceUnchanged
                InstalledIntoProject = $false; Deployed = $false
                Note = 'DLL existence is not embedded-label or runtime validation.'
            }
            Write-JsonFile (Join-Path $runRoot 'summary.json') $summary
            $zipPath = $runRoot + '.zip'
            Compress-Archive -LiteralPath $runRoot -DestinationPath $zipPath -CompressionLevel Optimal -ErrorAction Stop
            Write-Output ('RGSLAI_LABEL_STAGE_ZIP=' + $zipPath)
            Write-Output ('RGSLAI_LABEL_STAGE_SHA256=' + (Get-Sha256 $zipPath))
        }
        catch {
            $result = 'EVIDENCE_WRITE_FAILED'
            Write-Output ('RGSLAI_LABEL_EVIDENCE_ERROR=' + $_.Exception.Message)
            Write-Output ('RGSLAI_LABEL_LOGS_RETAINED_AT=' + $runRoot)
        }
    }
    Write-Output ('RGSLAI_LABEL_COMPILE_RESULT=' + $result)
    Write-Output 'RGSLAI_LABEL_PROJECT_RESOURCES_INSTALLED=NO'
    Write-Output 'RGSLAI_LABEL_DEPLOYED=NO'
}
