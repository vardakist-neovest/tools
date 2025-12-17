# Script to find duplicate ProfilerPath keys in production config files

# Define the array of clients to scan
$clients = @(
    # "GSM-DEV-INST2.layeronesoftware.com",
    # "GSM-DEV.layeronesoftware.com"
    # "client001-uat.layeronesoftware.com",
    # "client003.layeronesoftware.com",
    # "client003-uat.layeronesoftware.com",
    # "client007.layeronesoftware.com",
    # "client007-uat.layeronesoftware.com",
    # "client009.layeronesoftware.com",
    # "client011.layeronesoftware.com",
    # "client012.layeronesoftware.com",
    "client014.layeronesoftware.com",
    "client014-uat.layeronesoftware.com"
    # "client018.layeronesoftware.com",
    # "client020.layeronesoftware.com",
    # "client021.layeronesoftware.com",
    # "client021-uat.layeronesoftware.com",
    # "client023.layeronesoftware.com",
    # "client024.layeronesoftware.com",
    # "client024-uat.layeronesoftware.com",
    # "client024-qa.layeronesoftware.com",
    # "client025.layeronesoftware.com",
    # "client025-uat.layeronesoftware.com",
    # "client026.layeronesoftware.com",
    # "client027.layeronesoftware.com",
    # "client028.layeronesoftware.com",
    # "client029.layeronesoftware.com",
    # "client030.layeronesoftware.com",
    # "client031.layeronesoftware.com",
    # "client031-uat.layeronesoftware.com",
    # "client032.layeronesoftware.com",
    # "client033.layeronesoftware.com",
    # "client034.layeronesoftware.com",
    # "client035.layeronesoftware.com",
    # "client036.layeronesoftware.com",
    # "client037.layeronesoftware.com",
    # "client038.layeronesoftware.com",
    # "client039.layeronesoftware.com",
    # "client040.layeronesoftware.com",
    # "client041.layeronesoftware.com",
    # "client042.layeronesoftware.com",
    # "client043.layeronesoftware.com",
    # "client044.layeronesoftware.com",
    # "client045.layeronesoftware.com",
    # "client046.layeronesoftware.com",
    # "client047.layeronesoftware.com",
    # "client048.layeronesoftware.com",
    # "client049.layeronesoftware.com",
    # "client050.layeronesoftware.com",
    # "client051.layeronesoftware.com",
    # "client052.layeronesoftware.com",
    # "client053.layeronesoftware.com",
    # "client054.layeronesoftware.com"
)

$pattern = "*.config"

Write-Host "Scanning for duplicate ProfilerPath keys across multiple clients" -ForegroundColor Cyan
Write-Host ("=" * 80 -join "")

$results = @()

# First, scan local directory
$localPath = Join-Path $env:USERPROFILE "source\repos\POneNew"
Write-Host "`n--- Scanning LOCAL: $localPath ---" -ForegroundColor Cyan

if (Test-Path $localPath) {
    # Find all .exe.config files recursively in local directory
    $configFiles = Get-ChildItem -Path $localPath -Filter $pattern -Recurse -ErrorAction SilentlyContinue
    
    foreach ($configFile in $configFiles) {
        try {
            # Read the file content
            $content = Get-Content -Path $configFile.FullName -Raw
            
            # Skip if content is null or empty
            if ([string]::IsNullOrWhiteSpace($content)) {
                continue
            }
            
            # Count occurrences of ProfilerPath key
            $profilerMatches = [regex]::Matches($content, '<add\s+key\s*=\s*[''"]ProfilerPath[''"]', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            if ($profilerMatches.Count -gt 1) {
                Write-Host "`nFOUND DUPLICATE in: $($configFile.FullName)" -ForegroundColor Yellow
                Write-Host "  Count: $($profilerMatches.Count) occurrences" -ForegroundColor Red
                
                # Extract the lines with ProfilerPath
                $lines = Get-Content -Path $configFile.FullName
                $lineNumber = 0
                foreach ($line in $lines) {
                    $lineNumber++
                    if ($line -match 'add\s+key\s*=\s*[''"]ProfilerPath[''"]') {
                        Write-Host "  Line $lineNumber : $($line.Trim())" -ForegroundColor White
                    }
                }
                
                $results += [PSCustomObject]@{
                    Client = "LOCAL"
                    FilePath = $configFile.FullName
                    Count = $profilerMatches.Count
                }
            }
        }
        catch {
            Write-Host "Error reading $($configFile.FullName): $_" -ForegroundColor Red
        }
    }
} else {
    Write-Host "  Local path not found: $localPath" -ForegroundColor Red
}

# Loop through each client
foreach ($client in $clients) {
    $clientShortName = $client.Split('.')[0].ToUpper()  # Extract client### from FQDN
    $basePath = "\\$client\Services"
    
    Write-Host "`n--- Scanning $client ($basePath) ---" -ForegroundColor Magenta
    
    # Try primary path first
    if (-not (Test-Path $basePath)) {
        # Try alternate path pattern for some clients (e.g., client035 -> CL035)
        $altClientName = $clientShortName -replace 'CLIENT', 'CL'
        $altBasePath = "\\$client\Services$altClientName"
        
        if (Test-Path $altBasePath) {
            Write-Host "  Using alternate path: $altBasePath" -ForegroundColor Yellow
            $basePath = $altBasePath
        } else {
            Write-Host "  Path not accessible: $basePath (also tried $altBasePath)" -ForegroundColor Red
            continue
        }
    }
    
    # Get all subdirectories in Services folder
    $serviceFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue
    
    foreach ($folder in $serviceFolders) {
        # Find all .exe.config files in each service folder
        $configFiles = Get-ChildItem -Path $folder.FullName -Filter $pattern -ErrorAction SilentlyContinue
        
        foreach ($configFile in $configFiles) {
            try {
                # Read the file content
                $content = Get-Content -Path $configFile.FullName -Raw
                
                # Skip if content is null or empty
                if ([string]::IsNullOrWhiteSpace($content)) {
                    continue
                }
                
                # Count occurrences of ProfilerPath key
                $profilerMatches = [regex]::Matches($content, '<add\s+key\s*=\s*[''"]ProfilerPath[''"]', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                
                if ($profilerMatches.Count -gt 1) {
                    Write-Host "`nFOUND DUPLICATE in: $($configFile.FullName)" -ForegroundColor Yellow
                    Write-Host "  Count: $($profilerMatches.Count) occurrences" -ForegroundColor Red
                    
                    # Extract the lines with ProfilerPath
                    $lines = Get-Content -Path $configFile.FullName
                    $lineNumber = 0
                    foreach ($line in $lines) {
                        $lineNumber++
                        if ($line -match 'add\s+key\s*=\s*[''"]ProfilerPath[''"]') {
                            Write-Host "  Line $lineNumber : $($line.Trim())" -ForegroundColor White
                        }
                    }
                    
                    $results += [PSCustomObject]@{
                        Client = $client
                        FilePath = $configFile.FullName
                        Count = $profilerMatches.Count
                    }
                }
            }
            catch {
                Write-Host "Error reading $($configFile.FullName): $_" -ForegroundColor Red
            }
        }
    }
}

Write-Host "`n" + ("=" * 80)
Write-Host "`nSummary:" -ForegroundColor Cyan
if ($results.Count -eq 0) {
    Write-Host "No duplicate ProfilerPath keys found." -ForegroundColor Green
} else {
    Write-Host "Found $($results.Count) config file(s) with duplicate ProfilerPath keys:" -ForegroundColor Yellow
    Write-Host ""
    $results | Format-Table -Property Client, @{Label="Service/Config"; Expression={Split-Path $_.FilePath -Leaf}}, Count -AutoSize
    
    Write-Host "`nDetailed Paths:" -ForegroundColor Cyan
    $results | Format-Table -Property Client, FilePath, Count -AutoSize -Wrap
}

# Export results to CSV if any duplicates found
if ($results.Count -gt 0) {
    $csvPath = Join-Path $PSScriptRoot "duplicate-profiler-keys-results.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nResults exported to: $csvPath" -ForegroundColor Green
}
