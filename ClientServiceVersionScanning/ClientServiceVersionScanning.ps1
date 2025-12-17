
# Define the array of clients
# $clients = @("client042", "client044", "client014", "client014-UAT")
$clients = @("client014", "client014-UAT")
$instances = @("Portfolio3")

# Define the common path structure
$basePath = "\\{0}.layeronesoftware.com\Services\{1}\version.txt"

# Loop through each client and display the version.txt contents
foreach ($client in $clients) {
    foreach ($instance in $instances) {
        Write-Host "=== Versions for $client ==="
        
        $filePath = $basePath -f $client, $instance
        Write-Host "=== Path: $filePath ==="

        if (Test-Path $filePath) {
            Get-Content -Path $filePath
        } else {
            # Try alternate path pattern for some clients (e.g., client035 -> ServicesCL035)
            $clientShortName = $client.ToUpper()
            $altClientName = $clientShortName -replace 'CLIENT', 'CL'
            $altBasePath = "\\{0}.layeronesoftware.com\Services{1}\{2}\version.txt"
            $altFilePath = $altBasePath -f $client, $altClientName, $instance
            
            if (Test-Path $altFilePath) {
                Write-Host "Using alternate path: $altFilePath" -ForegroundColor Yellow
                Get-Content -Path $altFilePath
            } else {
                Write-Host "File not found for $client (tried both standard and alternate paths)" -ForegroundColor Red
            }
        }

        Write-Host "`n"  # Add a blank line for readability
    }
}
 