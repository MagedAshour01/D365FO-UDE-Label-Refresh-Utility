#requires -Version 5.1
<#
.SYNOPSIS
    Interactive D365 F&O label compiler for UDE/dev environments.

.DESCRIPTION
    Prompts for the key paths and modules, then compiles label resources with
    the selected LabelC.exe into a NEW staging folder.

    Safety behavior:
    - Does NOT modify label source XML/TXT files.
    - Does NOT copy DLLs back into the project automatically.
    - Does NOT deploy, build the full model, synchronize DB, restart services,
      or change PowerShell execution policy.
    - Records compiler identity, input hashes, logs, output hashes, and summary.

.NOTES
    Designed for Windows PowerShell 5.1.
#>

[CmdletBinding()]
param(
    [string]$MetadataRoot,
    [string]$ReferencePackagesRoot,
    [string[]]$Modules,
    [string]$Language = 'en-US',
    [string]$CompilerPath,
    [string]$OutputParent,
    [ValidateRange(1,120)]
    [int]$CompilerTimeoutMinutes = 10
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Read-Value {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [string]$Default
    )

    if ([string]::IsNullOrWhiteSpace($Default)) {
        do {
            $value = Read-Host $Prompt
        } while ([string]::IsNullOrWhiteSpace($value))
        return $value.Trim()
    }

    $value = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }
    return $value.Trim()
}

function Get-AbsoluteLocalPath {
    param([Parameter(Mandatory=$true)][string]$Value)

    if ($Value -notmatch '^[A-Za-z]:[\\/]' -or
        $Value -match '["\r\n]' -or
        $Value -match '[*?]') {
        throw "Use an absolute local drive path without wildcards or quotes: $Value"
    }

    return [IO.Path]::GetFullPath($Value).TrimEnd([char[]]@('\','/'))
}

function Get-Sha256 {
    param([Parameter(Mandatory=$true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][object]$Value
    )

    $json = ConvertTo-Json -InputObject $Value -Depth 12
    [IO.File]::WriteAllText($Path, $json, $utf8)
}

function Get-LabelDescriptorPath {
    param(
        [string]$MetadataRoot,
        [string]$Module,
        [string]$Language
    )

    return Join-Path $MetadataRoot (
        '{0}\{0}\AxLabelFile\{0}_{1}.xml' -f $Module, $Language
    )
}

function Get-LabelResourceTextPath {
    param(
        [string]$MetadataRoot,
        [string]$Module,
        [string]$Language
    )

    return Join-Path $MetadataRoot (
        '{0}\{0}\AxLabelFile\LabelResources\{1}\{0}.{1}.label.txt' -f $Module, $Language
    )
}

function Get-CandidateModules {
    param(
        [string]$MetadataRoot,
        [string]$Language
    )

    $found = @()

    foreach ($packageDir in @(Get-ChildItem -LiteralPath $MetadataRoot -Directory -ErrorAction Stop)) {
        $module = $packageDir.Name
        $descriptor = Get-LabelDescriptorPath -MetadataRoot $MetadataRoot -Module $module -Language $Language
        $resource = Get-LabelResourceTextPath -MetadataRoot $MetadataRoot -Module $module -Language $Language

        if ((Test-Path -LiteralPath $descriptor -PathType Leaf) -and
            (Test-Path -LiteralPath $resource -PathType Leaf)) {
            $found += $module
        }
    }

    return @($found | Sort-Object -Unique)
}

function Invoke-LabelCompiler {
    param(
        [string]$Module,
        [string]$MetadataRoot,
        [string]$CompilerPath,
        [string]$OutputPath,
        [string]$LogFolder,
        [int]$TimeoutMinutes
    )

    $stdoutPath = Join-Path $LogFolder 'stdout.txt'
    $stderrPath = Join-Path $LogFolder 'stderr.txt'
    $outLog = Join-Path $LogFolder 'labelc.log'
    $errLog = Join-Path $LogFolder 'labelc.err.xml'

    $arguments = '-metadata="{0}" -modelmodule="{1}" -output="{2}" -outlog="{3}" -errlog="{4}"' -f `
        $MetadataRoot, $Module, $OutputPath, $outLog, $errLog

    [IO.File]::WriteAllText(
        (Join-Path $LogFolder 'invocation.txt'),
        ($CompilerPath + [Environment]::NewLine + $arguments),
        $utf8
    )

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
        if (-not $process.Start()) {
            throw "Label compiler could not be started for module: $Module"
        }

        $stdOutTask = $process.StandardOutput.ReadToEndAsync()
        $stdErrTask = $process.StandardError.ReadToEndAsync()
        $deadline = [DateTime]::UtcNow.AddMinutes($TimeoutMinutes)

        while (-not $process.WaitForExit(1000)) {
            Write-Progress `
                -Activity "Compile labels: $Module" `
                -Status "Writing to staging only; source is unchanged."

            if ([DateTime]::UtcNow -gt $deadline) {
                $timedOut = $true
                $process.Kill()
                $process.WaitForExit()
                break
            }
        }

        $process.WaitForExit()

        [IO.File]::WriteAllText(
            $stdoutPath,
            $stdOutTask.GetAwaiter().GetResult(),
            $utf8
        )
        [IO.File]::WriteAllText(
            $stderrPath,
            $stdErrTask.GetAwaiter().GetResult(),
            $utf8
        )

        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            TimedOut = $timedOut
        }
    }
    finally {
        Write-Progress -Activity "Compile labels: $Module" -Completed
        $process.Dispose()
    }
}

$runRoot = $null
$result = 'NOT_STARTED'
$failureMessage = ''
$sourceBefore = @()
$moduleResults = @()
$compilerIdentity = $null
$writeOutputZip = $false

try {
    if ($env:OS -ne 'Windows_NT') {
        throw 'Run this script on the Windows D365 development/UDE machine.'
    }

    Write-Host ''
    Write-Host '=== D365 F&O Interactive Label Compiler ==='
    Write-Host 'Staging only - no automatic installation or deployment.'
    Write-Host ''

    if ([string]::IsNullOrWhiteSpace($MetadataRoot)) {
        $MetadataRoot = Read-Value -Prompt 'Metadata root' -Default ''
    }
    $MetadataRoot = Get-AbsoluteLocalPath $MetadataRoot

    if (-not (Test-Path -LiteralPath $MetadataRoot -PathType Container)) {
        throw "Metadata root not found: $MetadataRoot"
    }

    if ([string]::IsNullOrWhiteSpace($ReferencePackagesRoot)) {
        $ReferencePackagesRoot = Read-Value `
            -Prompt 'Reference PackagesLocalDirectory / package root' `
            -Default ''
    }
    $ReferencePackagesRoot = Get-AbsoluteLocalPath $ReferencePackagesRoot

    if (-not (Test-Path -LiteralPath $ReferencePackagesRoot -PathType Container)) {
        throw "Reference package root not found: $ReferencePackagesRoot"
    }

    if ([string]::IsNullOrWhiteSpace($Language)) {
        $Language = 'en-US'
    }

    Write-Host ''
    Write-Host "Scanning label modules for language: $Language ..."
    $candidates = @(Get-CandidateModules -MetadataRoot $MetadataRoot -Language $Language)

    if ($candidates.Count -gt 0) {
        Write-Host ''
        Write-Host 'Detected candidate modules:'
        for ($i = 0; $i -lt $candidates.Count; $i++) {
            Write-Host ("  [{0}] {1}" -f ($i + 1), $candidates[$i])
        }
    }
    else {
        Write-Warning "No modules with both descriptor and label TXT were auto-detected for $Language."
    }

    if ($null -eq $Modules -or $Modules.Count -eq 0) {
        Write-Host ''
        $moduleInput = Read-Host 'Enter module names separated by commas, or type ALL to use all detected modules'

        if ($moduleInput.Trim().ToUpperInvariant() -eq 'ALL') {
            if ($candidates.Count -eq 0) {
                throw 'ALL was selected, but no candidate modules were detected.'
            }
            $Modules = $candidates
        }
        else {
            $Modules = @(
                $moduleInput.Split(',') |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
            )
        }
    }

    if ($Modules.Count -eq 0) {
        throw 'At least one module is required.'
    }

    if ([string]::IsNullOrWhiteSpace($CompilerPath)) {
        $defaultCompiler = Join-Path $ReferencePackagesRoot 'bin\LabelC.exe'
        $CompilerPath = Read-Value -Prompt 'LabelC.exe path' -Default $defaultCompiler
    }
    $CompilerPath = Get-AbsoluteLocalPath $CompilerPath

    if (-not (Test-Path -LiteralPath $CompilerPath -PathType Leaf)) {
        throw "LabelC.exe was not found at: $CompilerPath"
    }

    if ([IO.Path]::GetFileName($CompilerPath) -ine 'LabelC.exe') {
        throw 'CompilerPath must point to LabelC.exe.'
    }

    if ([string]::IsNullOrWhiteSpace($OutputParent)) {
        $defaultOutput = 'D:\Temp'
        $OutputParent = Read-Value -Prompt 'Staging output parent' -Default $defaultOutput
    }
    $OutputParent = Get-AbsoluteLocalPath $OutputParent

    if (-not (Test-Path -LiteralPath $OutputParent -PathType Container)) {
        $create = Read-Host "Output folder does not exist. Create it? [Y/N]"
        if ($create.Trim().ToUpperInvariant() -ne 'Y') {
            throw "Output folder does not exist: $OutputParent"
        }
        New-Item -Path $OutputParent -ItemType Directory -Force | Out-Null
    }

    foreach ($protectedRoot in @($MetadataRoot, $ReferencePackagesRoot)) {
        if ($OutputParent.Equals($protectedRoot, [StringComparison]::OrdinalIgnoreCase) -or
            $OutputParent.StartsWith(($protectedRoot + '\'), [StringComparison]::OrdinalIgnoreCase)) {
            throw 'OutputParent must be outside MetadataRoot and ReferencePackagesRoot.'
        }
    }

    $compilerFile = Get-Item -LiteralPath $CompilerPath
    $compilerIdentity = [pscustomobject]@{
        Path = $compilerFile.FullName
        FileVersion = $compilerFile.VersionInfo.FileVersion
        ProductVersion = $compilerFile.VersionInfo.ProductVersion
        SHA256 = Get-Sha256 $CompilerPath
        Note = 'Compiler selected explicitly or from active reference package root.'
    }

    Write-Host ''
    Write-Host 'Selected configuration:'
    Write-Host "  MetadataRoot           : $MetadataRoot"
    Write-Host "  ReferencePackagesRoot  : $ReferencePackagesRoot"
    Write-Host "  CompilerPath           : $CompilerPath"
    Write-Host "  Language               : $Language"
    Write-Host "  Modules                : $($Modules -join ', ')"
    Write-Host "  OutputParent           : $OutputParent"
    Write-Host "  TimeoutMinutes         : $CompilerTimeoutMinutes"
    Write-Host ''

    $confirm = Read-Host 'Proceed with label compilation to staging only? [Y/N]'
    if ($confirm.Trim().ToUpperInvariant() -ne 'Y') {
        throw 'Cancelled by user before compilation.'
    }

    foreach ($module in $Modules) {
        $descriptorPath = Get-LabelDescriptorPath -MetadataRoot $MetadataRoot -Module $module -Language $Language
        $resourcePath = Get-LabelResourceTextPath -MetadataRoot $MetadataRoot -Module $module -Language $Language

        foreach ($requiredPath in @($descriptorPath, $resourcePath)) {
            if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
                throw "Required label source missing for module '$module': $requiredPath"
            }
        }

        $descriptor = New-Object System.Xml.XmlDocument
        $descriptor.XmlResolver = $null
        $descriptor.Load($descriptorPath)

        $idNode = $descriptor.SelectSingleNode('/AxLabelFile/LabelFileId')
        if ($null -eq $idNode -or $idNode.InnerText -cne $module) {
            throw "Unexpected LabelFileId in: $descriptorPath"
        }

        $labelRoot = Split-Path -Parent $descriptorPath

        foreach ($sourceFile in @(Get-ChildItem -LiteralPath $labelRoot -File -Recurse)) {
            $sourceBefore += [pscustomobject]@{
                Module = $module
                Path = $sourceFile.FullName
                SHA256 = Get-Sha256 $sourceFile.FullName
            }
        }
    }

    $runName = 'D365_LabelCompile_' +
               (Get-Date -Format 'yyyyMMdd_HHmmss_fff') +
               '_' +
               [Guid]::NewGuid().ToString('N').Substring(0,8)

    $runRoot = Join-Path $OutputParent $runName
    New-Item -Path $runRoot -ItemType Directory -ErrorAction Stop | Out-Null
    $writeOutputZip = $true

    Write-Output ('D365_LABEL_STAGE_DIRECTORY=' + $runRoot)
    Write-Output ('D365_LABEL_COMPILER=' + $CompilerPath)

    Write-JsonFile (Join-Path $runRoot 'compiler.json') $compilerIdentity
    Write-JsonFile (Join-Path $runRoot 'source-before.json') $sourceBefore

    foreach ($module in $Modules) {
        $moduleStage = Join-Path $runRoot $module
        $outputPath = Join-Path $moduleStage 'Resources'
        $logsPath = Join-Path $moduleStage 'Logs'
        $evidencePath = Join-Path $moduleStage 'SourceEvidence'

        foreach ($newPath in @($moduleStage, $outputPath, $logsPath, $evidencePath)) {
            New-Item -Path $newPath -ItemType Directory -ErrorAction Stop | Out-Null
        }

        $descriptorPath = Get-LabelDescriptorPath -MetadataRoot $MetadataRoot -Module $module -Language $Language
        $resourcePath = Get-LabelResourceTextPath -MetadataRoot $MetadataRoot -Module $module -Language $Language

        Copy-Item -LiteralPath $descriptorPath -Destination $evidencePath
        Copy-Item -LiteralPath $resourcePath -Destination $evidencePath

        Write-Host ''
        Write-Host "Compiling: $module"

        $execution = Invoke-LabelCompiler `
            -Module $module `
            -MetadataRoot $MetadataRoot `
            -CompilerPath $CompilerPath `
            -OutputPath $outputPath `
            -LogFolder $logsPath `
            -TimeoutMinutes $CompilerTimeoutMinutes

        $dlls = @(
            Get-ChildItem -LiteralPath $outputPath -Recurse -File -Filter '*.dll' |
            Where-Object { $_.Length -gt 0 }
        )

        $outputInventory = @(
            foreach ($file in @(Get-ChildItem -LiteralPath $outputPath -Recurse -File)) {
                [pscustomobject]@{
                    Path = $file.FullName
                    Bytes = $file.Length
                    LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
                    SHA256 = Get-Sha256 $file.FullName
                }
            }
        )

        Write-JsonFile (Join-Path $moduleStage 'output-inventory.json') $outputInventory

        $moduleResults += [pscustomobject]@{
            Module = $module
            ExitCode = $execution.ExitCode
            TimedOut = $execution.TimedOut
            NonEmptyDllCount = $dlls.Count
            OutputPath = $outputPath
        }

        Write-Output ('D365_LABEL_EXIT_CODE_' + $module + '=' + $execution.ExitCode)
        Write-Output ('D365_LABEL_DLL_COUNT_' + $module + '=' + $dlls.Count)

        if ($execution.TimedOut) {
            throw "Label compiler timed out for $module. Review staged logs."
        }

        if ($execution.ExitCode -ne 0) {
            throw "Label compiler failed for $module. Review staged logs."
        }

        if ($dlls.Count -eq 0) {
            throw "Compiler returned exit code 0 but generated no non-empty DLLs for $module."
        }
    }

    $result = 'STAGED_DLLS_REQUIRE_REVIEW'
}
catch {
    $result = 'FAILED_OR_BLOCKED'
    $failureMessage = $_.Exception.Message
    Write-Output ('D365_LABEL_ERROR=' + $failureMessage)
}
finally {
    $sourceUnchanged = $true

    foreach ($source in $sourceBefore) {
        try {
            if ((Get-Sha256 $source.Path) -cne $source.SHA256) {
                $sourceUnchanged = $false
            }
        }
        catch {
            $sourceUnchanged = $false
        }
    }

    if (-not $sourceUnchanged) {
        $result = 'SOURCE_CHANGED_OR_UNREADABLE_STOP'
        $failureMessage += ' Source hashes changed or became unreadable. Do not install staged outputs.'
    }

    if ($writeOutputZip -and -not [string]::IsNullOrWhiteSpace($runRoot)) {
        try {
            $summary = [pscustomobject]@{
                Result = $result
                Error = $failureMessage
                CheckUtc = [DateTime]::UtcNow.ToString('o')
                MetadataRoot = $MetadataRoot
                ReferencePackagesRoot = $ReferencePackagesRoot
                Language = $Language
                Compiler = $compilerIdentity
                Modules = $moduleResults
                ExistingLabelSourceFilesUnchanged = $sourceUnchanged
                InstalledIntoProject = $false
                Deployed = $false
                Note = 'Generated DLL existence does not prove runtime label resolution.'
            }

            Write-JsonFile (Join-Path $runRoot 'summary.json') $summary

            $zipPath = $runRoot + '.zip'
            Compress-Archive `
                -LiteralPath $runRoot `
                -DestinationPath $zipPath `
                -CompressionLevel Optimal `
                -ErrorAction Stop

            Write-Output ('D365_LABEL_STAGE_ZIP=' + $zipPath)
            Write-Output ('D365_LABEL_STAGE_SHA256=' + (Get-Sha256 $zipPath))
        }
        catch {
            $result = 'EVIDENCE_WRITE_FAILED'
            Write-Output ('D365_LABEL_EVIDENCE_ERROR=' + $_.Exception.Message)
            Write-Output ('D365_LABEL_LOGS_RETAINED_AT=' + $runRoot)
        }
    }

    Write-Host ''
    Write-Output ('D365_LABEL_COMPILE_RESULT=' + $result)
    Write-Output 'D365_LABEL_PROJECT_RESOURCES_INSTALLED=NO'
    Write-Output 'D365_LABEL_DEPLOYED=NO'
}
