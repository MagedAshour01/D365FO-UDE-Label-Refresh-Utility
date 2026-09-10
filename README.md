# D365 F&O UDE Label Refresh Utility

Interactive PowerShell utility for recompiling
D365 F&O Label Resource DLLs using LabelC.exe.

## Problem
Label metadata is updated but the changes are not reflected in UDE.

## What the utility does
- Validates MetadataRoot
- Validates PackagesLocalDirectory
- Detects label modules
- Resolves LabelC.exe
- Compiles into isolated staging
- Captures logs and SHA256 hashes
- Leaves source files unchanged
- No automatic deployment or overwrite

## Requirements
- Windows PowerShell 5.1
- D365 F&O UDE / development environment
- Valid MetadataRoot
- Matching PackagesLocalDirectory

## Usage
Run:
.\D365FO_UDE_Label_Refresh_Utility.ps1

Then follow the interactive prompts.
