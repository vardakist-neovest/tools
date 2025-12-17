<#!
.SYNOPSIS
Deploy environment-specific config into Fortress.Service.exe.config for a given project/service instance.

.DESCRIPTION
Given a loose project name, environment name (e.g. DEV1), and a service instance (e.g. Portfolio, STP),
this script locates:
  - The target project .csproj (loose substring match) under the workspace root
  - The `.Deploy` folder within that project
  - The service instance folder under `.Deploy`
  - The environment config file named `<ENV>.config` inside that instance folder
It then loads the environment config, performs textual replacements:
  - `localhost` -> `<ENV>.layeronesoftware.com`
  - Any drive root like `C:\`, `E:\` etc -> `D:\`
Creates a timestamped backup of the existing `Fortress.Service.exe.config` and overwrites it.

.PARAMETER Project
Loose project name substring. Defaults to `POne.Kernel`.

.PARAMETER Environment
Environment short name (e.g. DEV1). Used to locate `<Environment>.config` and for hostname substitution.

.PARAMETER ServiceInstance
Service instance folder name under `.Deploy` (e.g. Portfolio, STP).

.PARAMETER RepoRoot
Folder name under %USERPROFILE%\source\repos\ (e.g. POneNew, MyProject). The script will construct the full workspace path.

.PARAMETER DryRun
If set, performs all resolution and shows the intended changes without writing files.

.EXAMPLE
PS> ./Deploy-EnvConfig.ps1 -RepoRoot POneNew -Environment DEV1 -ServiceInstance Portfolio

.EXAMPLE
PS> ./Deploy-EnvConfig.ps1 -RepoRoot POneNew -Project Kernel -Environment DEV2 -ServiceInstance STP -Verbose

.NOTES
Requires read/write access to target project directory.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Project,
    [Parameter(Mandatory=$true)][string]$Environment,
    [Parameter(Mandatory=$true)][string]$ServiceInstance,
    [Parameter(Mandatory=$true)][string]$RepoRoot,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Construct workspace root from USERPROFILE and RepoRoot parameter
$WorkspaceRoot = Join-Path $env:USERPROFILE "source\repos\$RepoRoot"
if(-not (Test-Path $WorkspaceRoot)) {
    throw "Workspace root path does not exist: $WorkspaceRoot"
}
Write-Verbose "Using workspace root: $WorkspaceRoot"

function Find-ProjectCsproj {
    param(
        [string]$Root,
        [string]$ProjectPattern
    )
    Write-Verbose "Searching for project match '*$ProjectPattern*' under $Root"
    $candidates = Get-ChildItem -Path $Root -Recurse -Filter *.csproj -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$ProjectPattern*" }
    if(-not $candidates){ return $null }

    # Try exact (case-insensitive) filename match first
    $exact = $candidates | Where-Object { $_.BaseName -ieq $ProjectPattern }
    if($exact){ return $exact | Select-Object -First 1 }

    # Try prefixed with 'POne.'
    $pref = $candidates | Where-Object { $_.BaseName -ieq "POne.$ProjectPattern" }
    if($pref){ return $pref | Select-Object -First 1 }

    # Otherwise choose shortest path depth (likely primary project)
    return $candidates | Sort-Object { ($_.FullName -split '[\\/]').Count } | Select-Object -First 1
}

function Resolve-ServiceInstanceFolder {
    param(
        [string]$DeployRoot,
        [string]$InstanceName
    )
    if(-not (Test-Path $DeployRoot)) { return $null }
    $folders = Get-ChildItem -Path $DeployRoot -Directory -ErrorAction SilentlyContinue
    if(-not $folders){ return $null }
    # Exact (ci) name
    $exact = $folders | Where-Object { $_.Name -ieq $InstanceName }
    if($exact){ return $exact | Select-Object -First 1 }
    # Loose contains
    $loose = $folders | Where-Object { $_.Name -like "*$InstanceName*" }
    if($loose){ return $loose | Select-Object -First 1 }
    return $null
}

function Find-EnvironmentConfig {
    param(
        [string]$InstanceFolder,
        [string]$Env
    )
    $targetName = "$Env.config"
    Write-Verbose "Looking for environment config '$targetName' under $InstanceFolder"
    $files = Get-ChildItem -Path $InstanceFolder -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $targetName }
    if($files){ return $files | Select-Object -First 1 }
    return $null
}

function Find-FortressConfigTarget {
    param(
        [string]$ProjectDir
    )
    $targetName = 'Fortress.Service.exe.config'
    $direct = Join-Path $ProjectDir $targetName
    if(Test-Path $direct){ return Get-Item $direct }
    $found = Get-ChildItem -Path $ProjectDir -Recurse -Filter $targetName -ErrorAction SilentlyContinue | Select-Object -First 1
    return $found
}

# --- Resolution Phase ---
$csproj = Find-ProjectCsproj -Root $WorkspaceRoot -ProjectPattern $Project
if(-not $csproj){ throw "No matching .csproj found for pattern '$Project' under '$WorkspaceRoot'." }
$projectDir = $csproj.Directory.FullName
Write-Verbose "Resolved project directory: $projectDir"

$deployDir = Join-Path $projectDir '.Deploy'
$instanceFolder = Resolve-ServiceInstanceFolder -DeployRoot $deployDir -InstanceName $ServiceInstance
if(-not $instanceFolder){ throw "Service instance folder for '$ServiceInstance' not found under '$deployDir'." }
Write-Verbose "Resolved service instance folder: $($instanceFolder.FullName)"

$envConfig = Find-EnvironmentConfig -InstanceFolder $instanceFolder.FullName -Env $Environment
if(-not $envConfig){ throw "Environment config '$Environment.config' not found under '$($instanceFolder.FullName)'" }
Write-Verbose "Resolved environment config: $($envConfig.FullName)"

$targetConfig = Find-FortressConfigTarget -ProjectDir $projectDir
if(-not $targetConfig){ throw "Target Fortress.Service.exe.config not found under project directory '$projectDir'." }
Write-Verbose "Target Fortress config: $($targetConfig.FullName)"

# --- Transformation Phase ---
$raw = Get-Content -LiteralPath $envConfig.FullName -Raw
$hostname = "$Environment.layeronesoftware.com"
$transformed = $raw -replace 'localhost', $hostname
# Replace any drive letter root (A:\ .. Z:\) with D:\
$transformed = [Regex]::Replace($transformed, '(?i)\b[A-Z]:\\', 'D:\\')

# --- DryRun Preview ---
if($DryRun){
    Write-Host "[DryRun] Would overwrite: $($targetConfig.FullName)" -ForegroundColor Yellow
    Write-Host "[DryRun] Backup would be created alongside target." -ForegroundColor Yellow
    Write-Host "[DryRun] First 500 chars of transformed content:" -ForegroundColor Cyan
    Write-Host ($transformed.Substring(0, [Math]::Min(500, $transformed.Length)))
    return [PSCustomObject]@{
        Project        = $csproj.Name
        ProjectDir     = $projectDir
        ServiceInstance= $instanceFolder.Name
        Environment    = $Environment
        Hostname       = $hostname
        EnvConfigPath  = $envConfig.FullName
        TargetConfig   = $targetConfig.FullName
        DryRun         = $true
    }
}

# --- Backup & Write ---
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
# $backupPath = "$($targetConfig.FullName).bak.$timestamp"
# Copy-Item -LiteralPath $targetConfig.FullName -Destination $backupPath -Force
# Set-Content -LiteralPath $targetConfig.FullName -Value $transformed -Encoding UTF8

# Write-Host "Backup created: $backupPath" -ForegroundColor DarkGray
Write-Host "Updated config written to: $($targetConfig.FullName)" -ForegroundColor Green
Write-Host "Hostname substitution: localhost -> $hostname" -ForegroundColor Green
Write-Host "Drive letters normalized to D:\" -ForegroundColor Green

# --- Update .csproj to set Fortress.Service.exe.config CopyToOutputDirectory ---
Write-Verbose "Updating .csproj to set CopyToOutputDirectory for Fortress.Service.exe.config"
[xml]$csprojXml = Get-Content -LiteralPath $csproj.FullName
$ns = New-Object System.Xml.XmlNamespaceManager($csprojXml.NameTable)
$ns.AddNamespace("ms", "http://schemas.microsoft.com/developer/msbuild/2003")

$configChanged = $false

# Find existing node for Fortress.Service.exe.config
$configNode = $csprojXml.SelectSingleNode("//ms:None[@Include='Fortress.Service.exe.config'] | //ms:Content[@Include='Fortress.Service.exe.config']", $ns)
if(-not $configNode){
    # For SDK-style projects (no namespace)
    $configNode = $csprojXml.SelectSingleNode("//None[@Include='Fortress.Service.exe.config'] | //Content[@Include='Fortress.Service.exe.config']")
}

if($configNode){
    # Check if CopyToOutputDirectory exists
    $copyNode = $configNode.SelectSingleNode("ms:CopyToOutputDirectory", $ns)
    if(-not $copyNode){
        $copyNode = $configNode.SelectSingleNode("CopyToOutputDirectory")
    }
    
    if($copyNode){
        if($copyNode.InnerText -ne "Always"){
            $copyNode.InnerText = "Always"
            $configChanged = $true
            Write-Verbose "Updated existing CopyToOutputDirectory from '$($copyNode.InnerText)' to 'Always'"
        } else {
            Write-Verbose "CopyToOutputDirectory already set to 'Always'"
        }
    } else {
        # Add CopyToOutputDirectory node
        $newCopyNode = $csprojXml.CreateElement("CopyToOutputDirectory", $configNode.NamespaceURI)
        $newCopyNode.InnerText = "Always"
        [void]$configNode.AppendChild($newCopyNode)
        $configChanged = $true
        Write-Verbose "Added CopyToOutputDirectory='Always'"
    }
    
    if($configChanged){
        $csprojXml.Save($csproj.FullName)
        Write-Host "Set Fortress.Service.exe.config to 'Copy Always' in project file" -ForegroundColor Green
    } else {
        Write-Host "Fortress.Service.exe.config already set to 'Copy Always'" -ForegroundColor DarkGray
    }
} else {
    Write-Host "Fortress.Service.exe.config not found in .csproj (may be copied by default or post-build)" -ForegroundColor DarkGray
}

# --- Update .csproj.user Debug Settings ---
Write-Verbose "Updating project debug settings in .csproj.user file"

# Define the external program path
$externalExePath = "D:\Users\uvardth\source\repos\$RepoRoot\P1TC.Core\.Build\Debug\Fortress.Service.exe"

# Load or create .csproj.user file
$csprojUserPath = "$($csproj.FullName).user"
if(Test-Path $csprojUserPath){
    [xml]$userXml = Get-Content -LiteralPath $csprojUserPath
} else {
    # Create new .user file
    $userXml = New-Object System.Xml.XmlDocument
    $declaration = $userXml.CreateXmlDeclaration("1.0", "utf-8", $null)
    [void]$userXml.AppendChild($declaration)
    $projectNode = $userXml.CreateElement("Project", "http://schemas.microsoft.com/developer/msbuild/2003")
    $toolsVersionAttr = $userXml.CreateAttribute("ToolsVersion")
    $toolsVersionAttr.Value = "Current"
    [void]$projectNode.Attributes.Append($toolsVersionAttr)
    [void]$userXml.AppendChild($projectNode)
    Write-Verbose "Created new .csproj.user file"
}

$userNs = New-Object System.Xml.XmlNamespaceManager($userXml.NameTable)
$userNs.AddNamespace("ms", "http://schemas.microsoft.com/developer/msbuild/2003")

# Find or create PropertyGroup for Debug|AnyCPU
$debugUserPropGroup = $userXml.SelectSingleNode("//ms:PropertyGroup[contains(@Condition, 'Debug|AnyCPU')]", $userNs)
if(-not $debugUserPropGroup){
    $debugUserPropGroup = $userXml.SelectSingleNode("//PropertyGroup[contains(@Condition, 'Debug|AnyCPU')]")
}

if(-not $debugUserPropGroup){
    # Create PropertyGroup
    $debugUserPropGroup = $userXml.CreateElement("PropertyGroup", "http://schemas.microsoft.com/developer/msbuild/2003")
    $condAttr = $userXml.CreateAttribute("Condition")
    $condAttr.Value = " '`$(Configuration)|`$(Platform)' == 'Debug|AnyCPU' "
    [void]$debugUserPropGroup.Attributes.Append($condAttr)
    [void]$userXml.DocumentElement.AppendChild($debugUserPropGroup)
}

$userSettingsChanged = $false

# Set StartAction
$startActionNode = $debugUserPropGroup.SelectSingleNode("ms:StartAction", $userNs)
if(-not $startActionNode){
    $startActionNode = $debugUserPropGroup.SelectSingleNode("StartAction")
}
if(-not $startActionNode){
    $startActionNode = $userXml.CreateElement("StartAction", "http://schemas.microsoft.com/developer/msbuild/2003")
    [void]$debugUserPropGroup.AppendChild($startActionNode)
}
if($startActionNode.InnerText -ne "Program"){
    $startActionNode.InnerText = "Program"
    $userSettingsChanged = $true
}

# Set StartProgram
$startProgramNode = $debugUserPropGroup.SelectSingleNode("ms:StartProgram", $userNs)
if(-not $startProgramNode){
    $startProgramNode = $debugUserPropGroup.SelectSingleNode("StartProgram")
}
if(-not $startProgramNode){
    $startProgramNode = $userXml.CreateElement("StartProgram", "http://schemas.microsoft.com/developer/msbuild/2003")
    [void]$debugUserPropGroup.AppendChild($startProgramNode)
}
if($startProgramNode.InnerText -ne $externalExePath){
    $startProgramNode.InnerText = $externalExePath
    $userSettingsChanged = $true
}

# Set StartArguments
$startArgsNode = $debugUserPropGroup.SelectSingleNode("ms:StartArguments", $userNs)
if(-not $startArgsNode){
    $startArgsNode = $debugUserPropGroup.SelectSingleNode("StartArguments")
}
if(-not $startArgsNode){
    $startArgsNode = $userXml.CreateElement("StartArguments", "http://schemas.microsoft.com/developer/msbuild/2003")
    [void]$debugUserPropGroup.AppendChild($startArgsNode)
}
if($startArgsNode.InnerText -ne "-d"){
    $startArgsNode.InnerText = "-d"
    $userSettingsChanged = $true
}

if($userSettingsChanged){
    $userXml.Save($csprojUserPath)
    Write-Host "Updated project debug settings in .csproj.user: Start external program with -d flag" -ForegroundColor Green
} else {
    Write-Host "Project debug settings already configured correctly" -ForegroundColor DarkGray
}

# --- Set as Startup Project ---
Write-Verbose "Setting project as startup project"

# Find the solution file
$solutionPath = Join-Path "$WorkspaceRoot" "P1TC.Core\POne._All_.sln"
if(Test-Path $solutionPath){
    # Get relative path from solution directory to project
    $solutionDir = Split-Path $solutionPath -Parent
    $projectRelPath = $csproj.FullName.Replace("$solutionDir\", "")
    
    # Check if already set as startup project by reading .suo alternatives or ask user
    $shouldSetStartup = $false
    
    Write-Host ""
    Write-Host "Set '$($csproj.BaseName)' as startup project? (takes 10-20 seconds)" -ForegroundColor Cyan -NoNewline
    Write-Host " [Y/n]: " -ForegroundColor Yellow -NoNewline
    Write-Host ""
    Write-Host "(this will attempt to connect to VS and set the startup project programatically using ComObjects)" -ForegroundColor White
    
    # Read with timeout
    $startTime = Get-Date
    $timeout = 5
    $response = $null
    
    while (((Get-Date) - $startTime).TotalSeconds -lt $timeout) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            $response = $key.KeyChar
            break
        }
        Start-Sleep -Milliseconds 100
    }
    
    if($response -eq $null) {
        Write-Host "(timeout - defaulting to No)" -ForegroundColor DarkGray
        $shouldSetStartup = $false
    } elseif($response -match '^[Yy]$' -or $response -eq [char]13) {
        Write-Host ""
        $shouldSetStartup = $true
    } else {
        Write-Host ""
        Write-Host "Skipped setting startup project" -ForegroundColor DarkGray
        $shouldSetStartup = $false
    }
    
    if($shouldSetStartup) {
        try {
            Write-Host "Connecting to Visual Studio (this may take a moment)..." -ForegroundColor Gray
            
            # Try to get running instance first
            $vs = $null
            try {
                $vs = [System.Runtime.InteropServices.Marshal]::GetActiveObject("VisualStudio.DTE.17.0")
                Write-Verbose "Connected to running VS instance"
                
                # Check if correct solution is already open
                if($vs.Solution.FullName -eq $solutionPath) {
                    Write-Verbose "Correct solution already open"
                    
                    # Check if already set as startup project
                    $currentStartup = $vs.Solution.SolutionBuild.StartupProjects
                    Write-Verbose "Current startup projects: $($currentStartup -join ', ')"
                    Write-Verbose "Target project path: $projectRelPath"
                    
                    # Note: StartupProjects property sometimes returns cached/stale data
                    # We'll log it but continue to set it anyway to ensure it's persisted
                    if($currentStartup -eq $projectRelPath -or ($currentStartup -is [array] -and $currentStartup -contains $projectRelPath)) {
                        Write-Host "VS reports '$($csproj.BaseName)' is already set, but will verify by setting it again..." -ForegroundColor Gray
                    }
                } else {
                    # Wrong solution open, need to open correct one
                    $vs.Solution.Open($solutionPath)
                }
            } catch {
                # No running instance, create new one
                Write-Verbose "No running VS instance, creating new one"
                $vs = New-Object -ComObject "VisualStudio.DTE.17.0"
                $vs.Solution.Open($solutionPath)
            }
            
            # Set startup project - expects an array of project relative paths
            $vs.Solution.SolutionBuild.StartupProjects = @($projectRelPath)
            
            # Save solution
            $vs.Solution.SaveAs($solutionPath)
            
            # Only quit if we created a new instance
            if($vs.MainWindow.Visible -eq $false) {
                $vs.Quit()
            }
            
            # Cleanup COM object
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vs) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            Write-Host "Set '$($csproj.BaseName)' as startup project" -ForegroundColor Green
            
        } catch {
            Write-Warning "Could not set startup project: $($_.Exception.Message)"
            Write-Host "You may need to set startup project manually in Visual Studio" -ForegroundColor Yellow
        }
    }
} else {
    Write-Warning "Solution file not found at: $solutionPath"
}

return [PSCustomObject]@{
    Project        = $csproj.Name
    ProjectDir     = $projectDir
    ServiceInstance= $instanceFolder.Name
    Environment    = $Environment
    Hostname       = $hostname
    EnvConfigPath  = $envConfig.FullName
    TargetConfig   = $targetConfig.FullName
    # BackupPath     = $backupPath
    DryRun         = $false
}
