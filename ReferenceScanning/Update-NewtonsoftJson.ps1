param(
    [switch]$DryRun = $false,
    [string]$NewVersion = "13.0.3",
    [string]$RepositoryPath = (Get-Location).Path,
    [switch]$OnlyDiff = $false
)

<#
.SYNOPSIS
    Updates all Newtonsoft.Json references from version 13.0.0 to a newer version across the entire repository.

.DESCRIPTION
    This script searches for and updates:
    - Assembly references in .csproj files (Version=13.0.0.0)
    - Binding redirects in config files (newVersion="13.0.0.0")
    - PackageReference entries (if any)
    - packages.config entries (if needed)

.PARAMETER DryRun
    When specified, shows what would be changed without making actual modifications.

.PARAMETER NewVersion
    The new version to update to (default: 13.0.3). This should be the 3-part version number.

.PARAMETER RepositoryPath
    The root path of the repository to scan (default: current directory).

.PARAMETER OnlyDiff
    When used with -DryRun, shows only files that actually need updates (red entries in the table).
    Files that already match the target version are filtered out.

.EXAMPLE
    .\Update-NewtonsoftJson.ps1 -DryRun
    Shows what changes would be made without applying them.

.EXAMPLE
    .\Update-NewtonsoftJson.ps1 -DryRun -OnlyDiff
    Shows only files that need updates, filtering out files that already match the target version.

.EXAMPLE
    .\Update-NewtonsoftJson.ps1 -NewVersion "13.0.3"
    Updates all references to version 13.0.3.
#>

# Validate version format
if ($NewVersion -notmatch '^\d+\.\d+\.\d+$') {
    Write-Error "NewVersion must be in format X.Y.Z (e.g., 13.0.3)"
    exit 1
}

# Convert to 4-part version for assembly references
$NewAssemblyVersion = "$NewVersion.0"

Write-Host "Newtonsoft.Json Version Update Script" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Repository Path: $RepositoryPath" -ForegroundColor Yellow
Write-Host "Target Version: $NewVersion (Assembly: $NewAssemblyVersion)" -ForegroundColor Yellow
Write-Host "Dry Run Mode: $DryRun" -ForegroundColor Yellow
Write-Host ""

# Initialize counters
$changesCount = 0
$filesModified = @()
$fileVersions = @{}

function Write-Change {
    param($FilePath, $LineNumber, $OldText, $NewText, $ChangeType)
    
    $script:changesCount++
    $relativePath = $FilePath.Replace($RepositoryPath, "").TrimStart('\', '/')
    
    # Extract version from old text
    $versionMatch = $null
    $isBindingRedirect = $false
    
    if ($OldText -match '13\.0\.(\d+)(\.0)?') {
        $versionMatch = $matches[0]
        # Check if this is from a binding redirect
        $isBindingRedirect = $OldText -match 'bindingRedirect|newVersion'
    }
    # Also check for HintPath versions (e.g., Newtonsoft.Json.13.0.4\lib)
    elseif ($OldText -match 'Newtonsoft\.Json\.(13\.0\.(\d+))\\') {
        $versionMatch = $matches[1]  # Extract the 3-part version from HintPath
    }
    
    # Track file and version with source information
    if ($versionMatch) {
        if (-not $script:fileVersions.ContainsKey($relativePath)) {
            $script:fileVersions[$relativePath] = @()
        }
        
        # Create version info object with source type
        $versionInfo = @{
            Version = $versionMatch
            IsBindingRedirect = $isBindingRedirect
        }
        
        # Only add if we don't already have this exact version info
        $alreadyExists = $false
        foreach ($existing in $script:fileVersions[$relativePath]) {
            if ($existing.Version -eq $versionMatch -and $existing.IsBindingRedirect -eq $isBindingRedirect) {
                $alreadyExists = $true
                break
            }
        }
        
        if (-not $alreadyExists) {
            $script:fileVersions[$relativePath] += $versionInfo
        }
    }
    
    Write-Host "[$ChangeType] $relativePath" -ForegroundColor Green
    Write-Host "  Line $LineNumber" -ForegroundColor Gray
    Write-Host "  - $OldText" -ForegroundColor Red
    Write-Host "  + $NewText" -ForegroundColor Green
    Write-Host ""
    
    if ($script:filesModified -notcontains $FilePath) {
        $script:filesModified += $FilePath
    }
}

function Update-FileContent {
    param($FilePath, $Pattern, $Replacement, $ChangeType)
    
    if (-not (Test-Path $FilePath)) {
        return
    }
    
    $lines = Get-Content $FilePath
    $modified = $false
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match $Pattern) {
            $newLine = $line -replace $Pattern, $Replacement
            if ($line -ne $newLine) {
                Write-Change -FilePath $FilePath -LineNumber ($i + 1) -OldText $line.Trim() -NewText $newLine.Trim() -ChangeType $ChangeType
                
                if (-not $DryRun) {
                    $lines[$i] = $newLine
                    $modified = $true
                }
            }
        }
    }
    
    if ($modified -and -not $DryRun) {
        $lines | Set-Content $FilePath -Encoding UTF8
    }
}

function Scan-ExistingVersions {
    param($FilePath)
    
    if (-not (Test-Path $FilePath)) {
        return
    }
    
    $relativePath = $FilePath.Replace($RepositoryPath, "").TrimStart('\', '/')
    $content = Get-Content $FilePath -Raw
    
    # Look for HintPath versions that might be newer than assembly reference versions
    if ($content -match 'Newtonsoft\.Json\.(13\.0\.(\d+))\\') {
        $hintPathVersion = $matches[1]  # Extract the 3-part version
        
        if (-not $script:fileVersions.ContainsKey($relativePath)) {
            $script:fileVersions[$relativePath] = @()
        }
        
        # Create version info object
        $versionInfo = @{
            Version = $hintPathVersion
            IsBindingRedirect = $false
        }
        
        # Only add if we don't already have this version
        $alreadyExists = $false
        foreach ($existing in $script:fileVersions[$relativePath]) {
            if ($existing.Version -eq $hintPathVersion -and $existing.IsBindingRedirect -eq $false) {
                $alreadyExists = $true
                break
            }
        }
        
        if (-not $alreadyExists) {
            $script:fileVersions[$relativePath] += $versionInfo
        }
    }
}

function Show-FileVersionTable {
    if ($script:fileVersions.Count -eq 0) {
        Write-Host "No files with version information found." -ForegroundColor Yellow
        return
    }
    
    # Calculate column widths
    $maxFileLength = ($script:fileVersions.Keys | Measure-Object -Property Length -Maximum).Maximum
    $maxVersionLength = ($script:fileVersions.Values | ForEach-Object { $_ -join ", " } | Measure-Object -Property Length -Maximum).Maximum
    $targetVersionText = "$NewVersion ($NewAssemblyVersion)"
    
    $fileColWidth = [Math]::Max($maxFileLength, 4) + 2  # "File" header + padding
    $currentVersionColWidth = [Math]::Max($maxVersionLength, 15) + 2  # "Current Version(s)" header + padding
    $targetVersionColWidth = [Math]::Max($targetVersionText.Length, 14) + 2  # "Target Version" header + padding
    
    # Table header
    $separator = "+" + ("-" * $fileColWidth) + "+" + ("-" * $currentVersionColWidth) + "+" + ("-" * $targetVersionColWidth) + "+"
    Write-Host $separator -ForegroundColor Cyan
    Write-Host ("|" + " File".PadRight($fileColWidth) + "|" + " Curr Ver(s)".PadRight($currentVersionColWidth) + "|" + " Target Ver".PadRight($targetVersionColWidth) + "|") -ForegroundColor Cyan
    Write-Host $separator -ForegroundColor Cyan
    
    # Table rows
    $sortedFiles = $script:fileVersions.Keys | Sort-Object
    $filesToShow = @()
    
    # Pre-filter files if OnlyDiff is specified
    foreach ($file in $sortedFiles) {
        $showFile = $true
        
        if ($OnlyDiff -and $DryRun) {
            # Check if this file needs updates (has any meaningful version that doesn't match)
            $needsUpdate = $false
            $hasCorrectHintPath = $false
            
            # First, check if there's already a correct HintPath
            foreach ($versionInfo in $script:fileVersions[$file]) {
                if (-not $versionInfo.IsBindingRedirect -and $versionInfo.Version -eq $NewVersion) {
                    $hasCorrectHintPath = $true
                    break
                }
            }
            
            # If there's already a correct HintPath, don't show this file unless there are binding redirects that need updates
            if ($hasCorrectHintPath) {
                # Only check binding redirects for updates
                foreach ($versionInfo in $script:fileVersions[$file]) {
                    if ($versionInfo.IsBindingRedirect) {
                        $currentVer = $versionInfo.Version
                        
                        # Normalize both versions to 4-part format for comparison
                        $normalizedCurrent = $currentVer
                        if ($currentVer -match '^\d+\.\d+\.\d+$') { 
                            $normalizedCurrent = "$currentVer.0" 
                        }
                        
                        $normalizedTarget = $NewAssemblyVersion
                        
                        # For binding redirects, only check major version (13.0) match
                        $currentMajor = ""
                        $targetMajor = ""
                        if ($normalizedCurrent -match '^(\d+\.\d+)\.') {
                            $currentMajor = $matches[1]
                        }
                        if ($normalizedTarget -match '^(\d+\.\d+)\.') {
                            $targetMajor = $matches[1]
                        }
                        
                        # If binding redirect major version doesn't match, this file needs updates
                        if ($currentMajor -ne $targetMajor) {
                            $needsUpdate = $true
                            break
                        }
                    }
                }
            } else {
                # No correct HintPath, check all versions for mismatches
                foreach ($versionInfo in $script:fileVersions[$file]) {
                    $currentVer = $versionInfo.Version
                    $isBindingRedirect = $versionInfo.IsBindingRedirect
                    
                    # Normalize both versions to 4-part format for comparison
                    $normalizedCurrent = $currentVer
                    if ($currentVer -match '^\d+\.\d+\.\d+$') { 
                        $normalizedCurrent = "$currentVer.0" 
                    }
                    
                    $normalizedTarget = $NewAssemblyVersion
                    $versionMatches = $false
                    
                    # For binding redirects, only check major version (13.0) match
                    if ($isBindingRedirect) {
                        # Extract major version (13.0) from both current and target
                        $currentMajor = ""
                        $targetMajor = ""
                        if ($normalizedCurrent -match '^(\d+\.\d+)\.') {
                            $currentMajor = $matches[1]
                        }
                        if ($normalizedTarget -match '^(\d+\.\d+)\.') {
                            $targetMajor = $matches[1]
                        }
                        
                        $versionMatches = ($currentMajor -eq $targetMajor)  # Both should be 13.0
                    } else {
                        # For non-binding redirects, require exact match
                        $versionMatches = ($normalizedCurrent -eq $normalizedTarget)
                    }
                    
                    # If any version doesn't match, this file needs updates
                    if (-not $versionMatches) {
                        $needsUpdate = $true
                        break
                    }
                }
            }
            
            # Only show files that need updates
            $showFile = $needsUpdate
        }
        
        if ($showFile) {
            $filesToShow += $file
        }
    }
    
    # If OnlyDiff is specified and no files need updates, show message
    if ($OnlyDiff -and $filesToShow.Count -eq 0) {
        Write-Host "No files need updates - all versions already match target!" -ForegroundColor Green
        return
    }
    
    foreach ($file in $filesToShow) {
        # Extract just version strings for display
        $currentVersionStrings = @()
        foreach ($versionInfo in $script:fileVersions[$file]) {
            $versionStr = $versionInfo.Version
            if ($versionInfo.IsBindingRedirect) {
                $versionStr += " (BR)"  # Mark binding redirects
            }
            $currentVersionStrings += $versionStr
        }
        
        $currentVersions = $currentVersionStrings -join ", "
        $fileCell = " $file".PadRight($fileColWidth)
        $currentVersionCell = " $currentVersions".PadRight($currentVersionColWidth)
        $targetVersionCell = " $targetVersionText".PadRight($targetVersionColWidth)
        
        # Determine target version color based on whether versions match
        $targetVersionColor = "DarkGreen"  # Default to green
        if ($DryRun) {
            $hasMatch = $false
            $highestCurrentVersion = "0.0.0"
            
            foreach ($versionInfo in $script:fileVersions[$file]) {
                $currentVer = $versionInfo.Version
                $isBindingRedirect = $versionInfo.IsBindingRedirect
                
                # Normalize both versions to 4-part format for comparison
                $normalizedCurrent = $currentVer
                if ($currentVer -match '^\d+\.\d+\.\d+$') { 
                    $normalizedCurrent = "$currentVer.0" 
                }
                
                # Track the highest version found
                if ([System.Version]$normalizedCurrent -gt [System.Version]"$highestCurrentVersion.0") {
                    $highestCurrentVersion = $currentVer
                }
                
                $normalizedTarget = $NewAssemblyVersion
                
                # For binding redirects, only check major version (13.0) match
                if ($isBindingRedirect) {
                    # Extract major version (13.0) from both current and target
                    $currentMajor = ""
                    $targetMajor = ""
                    if ($normalizedCurrent -match '^(\d+\.\d+)\.') {
                        $currentMajor = $matches[1]
                    }
                    if ($normalizedTarget -match '^(\d+\.\d+)\.') {
                        $targetMajor = $matches[1]
                    }
                    
                    if ($currentMajor -eq $targetMajor) {  # Both should be 13.0
                        $hasMatch = $true
                        break
                    }
                } else {
                    # For non-binding redirects, require exact match
                    if ($normalizedCurrent -eq $normalizedTarget) {
                        $hasMatch = $true
                        break
                    }
                }
            }
            
            # If no exact match, check if the highest version (likely from HintPath) matches target
            if (-not $hasMatch) {
                $normalizedHighest = if ($highestCurrentVersion -match '^\d+\.\d+\.\d+$') { "$highestCurrentVersion.0" } else { $highestCurrentVersion }
                if ($normalizedHighest -eq $NewAssemblyVersion) {
                    $hasMatch = $true
                }
            }
            
            $targetVersionColor = if ($hasMatch) { "DarkGreen" } else { "Red" }
        }
        
        # Write each column with different colors
        Write-Host "|" -NoNewline -ForegroundColor Cyan
        Write-Host $fileCell -NoNewline -ForegroundColor White
        Write-Host "|" -NoNewline -ForegroundColor Cyan
        Write-Host $currentVersionCell -NoNewline -ForegroundColor Yellow
        Write-Host "|" -NoNewline -ForegroundColor Cyan
        Write-Host $targetVersionCell -NoNewline -ForegroundColor $targetVersionColor
        Write-Host "|" -ForegroundColor Cyan
    }
    
    # Table footer
    Write-Host $separator -ForegroundColor Cyan
    Write-Host ""
}

# 1. Update .csproj files - Assembly References
Write-Host "Scanning for .csproj files with Newtonsoft.Json assembly references..." -ForegroundColor Magenta

$csprojFiles = Get-ChildItem -Path $RepositoryPath -Recurse -Filter "*.csproj" | Where-Object { 
    $_.FullName -notlike "*\bin\*" -and $_.FullName -notlike "*\obj\*" 
}

foreach ($file in $csprojFiles) {
    # First scan for existing versions (including HintPath versions)
    Scan-ExistingVersions -FilePath $file.FullName
    
    # Pattern for assembly reference version
    $pattern = '(Newtonsoft\.Json, Version=)13\.0\.0\.0(, Culture=neutral, PublicKeyToken=30ad4fe6b2a6aeed)'
    $replacement = "`\${1}$NewAssemblyVersion`\${2}"
    Update-FileContent -FilePath $file.FullName -Pattern $pattern -Replacement $replacement -ChangeType "CSPROJ-ASSEMBLY-REF"
    
    # Pattern for HintPath references
    $hintPathPattern = '(\\packages\\Newtonsoft\.Json\.)13\.0\.0(\\)'
    $hintPathReplacement = "`\${1}$NewVersion`\${2}"
    Update-FileContent -FilePath $file.FullName -Pattern $hintPathPattern -Replacement $hintPathReplacement -ChangeType "CSPROJ-HINT-PATH"
}

# 2. Update config files - Binding Redirects
Write-Host "Scanning for config files with binding redirects..." -ForegroundColor Magenta

$configFiles = Get-ChildItem -Path $RepositoryPath -Recurse -Include "*.config", "web.config", "app.config" | Where-Object { 
    $_.FullName -notlike "*\bin\*" -and $_.FullName -notlike "*\obj\*" 
}

foreach ($file in $configFiles) {
    # Pattern for binding redirect newVersion
    $pattern = '(bindingRedirect oldVersion="[^"]*-?)13\.0\.0\.0(" newVersion=")13\.0\.0\.0(")'
    $replacement = "`${1}$NewAssemblyVersion`${2}$NewAssemblyVersion`${3}"
    Update-FileContent -FilePath $file.FullName -Pattern $pattern -Replacement $replacement -ChangeType "CONFIG-BINDING-REDIRECT"
    
    # Alternative pattern for simpler binding redirects
    $simplePattern = '(newVersion=")13\.0\.0\.0(")'
    $simpleReplacement = "`${1}$NewAssemblyVersion`${2}"
    Update-FileContent -FilePath $file.FullName -Pattern $simplePattern -Replacement $simpleReplacement -ChangeType "CONFIG-NEW-VERSION"
}

# 3. Update packages.config files
Write-Host "Scanning for packages.config files..." -ForegroundColor Magenta

$packagesConfigFiles = Get-ChildItem -Path $RepositoryPath -Recurse -Filter "packages.config" | Where-Object { 
    $_.FullName -notlike "*\bin\*" -and $_.FullName -notlike "*\obj\*" 
}

foreach ($file in $packagesConfigFiles) {
    # Only update if currently on 13.0.0
    $pattern = '(<package id="Newtonsoft\.Json" version=")13\.0\.0(" targetFramework="[^"]*" />)'
    $replacement = "`${1}$NewVersion`${2}"
    Update-FileContent -FilePath $file.FullName -Pattern $pattern -Replacement $replacement -ChangeType "PACKAGES-CONFIG"
}

# 4. Update PackageReference entries (if any exist)
Write-Host "Scanning for PackageReference entries..." -ForegroundColor Magenta

$projectFiles = Get-ChildItem -Path $RepositoryPath -Recurse -Include "*.csproj", "*.vbproj", "*.fsproj" | Where-Object { 
    $_.FullName -notlike "*\bin\*" -and $_.FullName -notlike "*\obj\*" 
}

foreach ($file in $projectFiles) {
    $pattern = '(<PackageReference Include="Newtonsoft\.Json" Version=")13\.0\.0(" />)'
    $replacement = "`${1}$NewVersion`${2}"
    Update-FileContent -FilePath $file.FullName -Pattern $pattern -Replacement $replacement -ChangeType "PACKAGE-REFERENCE"
}

# Summary
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "Total changes found: $changesCount" -ForegroundColor Yellow
Write-Host "Files that would be modified: $($filesModified.Count)" -ForegroundColor Yellow

if ($DryRun) {
    Write-Host ""
    if ($script:fileVersions.Count -gt 0) {
        Write-Host "Files and Versions Found:" -ForegroundColor Cyan
        Show-FileVersionTable
    }
    Write-Host "This was a DRY RUN - no files were actually modified." -ForegroundColor Yellow
    Write-Host "Run the script without -DryRun to apply these changes." -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "All changes have been applied successfully!" -ForegroundColor Green
    
    if ($filesModified.Count -gt 0) {
        Write-Host ""
        Write-Host "Modified files:" -ForegroundColor Cyan
        foreach ($file in $filesModified) {
            $relativePath = $file.Replace($RepositoryPath, "").TrimStart('\', '/')
            Write-Host "  $relativePath" -ForegroundColor Gray
        }
        
        Write-Host ""
        Write-Host "IMPORTANT NEXT STEPS:" -ForegroundColor Red
        Write-Host "1. Rebuild your solution to ensure all references are updated correctly" -ForegroundColor Yellow
        Write-Host "2. Update NuGet packages if necessary: Update-Package Newtonsoft.Json" -ForegroundColor Yellow
        Write-Host "3. Test your applications thoroughly" -ForegroundColor Yellow
        Write-Host "4. Commit your changes to version control" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Script completed." -ForegroundColor Cyan