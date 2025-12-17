Import-Module SqlServer

$EnvironmentsFolder = "D:\Users\uvardth\OneDrive - Neovest\Documents\Environments"
$DevFolder = Join-Path -Path $EnvironmentsFolder -ChildPath "Dev"
$ProdFolder = Join-Path -Path $EnvironmentsFolder -ChildPath "Prod"

function CreateSSMSEntry {
    param (
        [string]$ServerGroup,
        [string]$ServerName,
        [string]$ServerAddress
    )
    $OriginalLocation = Get-Location
    $BasePath = "SQLSERVER:\SQLRegistration\Database Engine Server Group"

    Set-Location $BasePath

    $GroupPath = Join-Path $BasePath $ServerGroup

    # Create group if it doesn't exist
    if (-not (Test-Path $GroupPath)) {
        Write-Host "Creating server group: $ServerGroup"
        New-Item -Path $BasePath -Name $ServerGroup | Out-Null
    }

    # Registration connection string
    $ConnectionString = "Server=$ServerAddress;Integrated Security=True;TrustServerCertificate=True;Database=FTBTrade;"

    # Register server if not already registered
    $RegistrationPath = Join-Path $GroupPath $ServerName
    $LogRegistration = $true;
    if (Test-Path $RegistrationPath) {
        Remove-Item $RegistrationPath
        $LogRegistration = $false;
    }
    New-Item -Path $GroupPath -Name $ServerName -ItemType Registration -Value $ConnectionString | Out-Null

    Set-Location $OriginalLocation

    if ($LogRegistration) {
        Write-Host "Created new server registration '$ServerName' in group '$ServerGroup' with address '$ServerAddress'"
    }
}

function CreateShortcut {
    param (
        [string]$ShortcutName,
        [string]$TargetPath,
        [string]$OutputFolder
    )

    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory | Out-Null
    }

    if (Test-Path (Join-Path -Path $OutputFolder -ChildPath ("{0}.lnk" -f $ShortcutName))) {
        return
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcutPath = Join-Path -Path $OutputFolder -ChildPath ("{0}.lnk" -f $ShortcutName)
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Save()

    Write-Host "Created shortcut '$ShortcutName' for '$TargetPath' in '$OutputFolder'"
}

function Add-Credential {
    param (
        [string]$Target,
        [string]$Username,
        [string]$Password
    )

    Write-Host "Adding credential for $Target..."
    $timeoutSeconds = 10
    $scriptBlock = {
        param($Target, $Username, $Password)
        cmdkey /add:$Target /user:$Username /pass:$Password
    }
    $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $Target, $Username, $Password
    $result = Wait-Job -Job $job -Timeout $timeoutSeconds
    if ($result -eq $null) {
        Write-Host "`u001b[31mTIMEOUT adding credential for $Target`u001b[0m"
        Stop-Job -Job $job | Out-Null
        Remove-Job -Job $job | Out-Null
    } else {
        Receive-Job -Job $job | Out-Null
        Write-Host "Credential added for $Target"
        Remove-Job -Job $job | Out-Null
    }
}

function Safe-Add-Credential {
    param (
        [string]$Target,
        [string]$Username,
        [string]$Password
    )

    $existing = cmdkey /list:$Target
    if (-not ($existing -match "\* NONE \*")) {
        return
    }

    Add-Credential -Target $Target -Username $Username -Password $Password
}

$CredentialsConstants = @{
    "layeronedev"    = @{ Username = "layeronedev\tvardakis";    Password = "Ath3nsUs3rAccount2025" }
    "convex"  = @{ Username = "convex\tvardakis";  Password = "Ath3nsUs3rAccount2025" }
    "layeronecloud" = @{ Username = "layeronecloud\tvardakis"; Password = "col%Sn3eze66EKW9E" }
    "client007"  = @{ Username = "client007\tvardakis"; Password = "{2b=>5KOu@S""o]NK" }
    "client008"  = @{ Username = "client008\user"; Password = "password" }
    "client009"  = @{ Username = "client009\user"; Password = "password" }
    "client011"  = @{ Username = "client011\user"; Password = "password" }
}

$DevEnvironments = @(
    @{ Name = "Dev1"; AppServer = "dev1.layeronesoftware.com"; SqlServer = "dev1.layeronesoftware.com,1433"; Credential = "layeronedev" }
    @{ Name = "Dev1-Inst2"; AppServer = "dev1.layeronesoftware.com"; SqlServer = "dev1.layeronesoftware.com,1435"; Credential = "layeronedev" }
    @{ Name = "Dev2"; AppServer = "dev2.layeronesoftware.com"; SqlServer = "dev2.layeronesoftware.com,1433"; Credential = "convex" }
    @{ Name = "Dev2-Inst2"; AppServer = "dev2.layeronesoftware.com"; SqlServer = "dev2.layeronesoftware.com,1435"; Credential = "convex" }
    @{ Name = "Dev3"; AppServer = "dev3.layeronesoftware.com"; SqlServer = "dev3.layeronesoftware.com,1433"; Credential = "layeronedev" }
    @{ Name = "Dev3-Inst2"; AppServer = "dev3.layeronesoftware.com"; SqlServer = "dev3.layeronesoftware.com,1435"; Credential = "layeronedev" }
    @{ Name = "Dev4"; AppServer = "dev4.layeronesoftware.com"; SqlServer = "dev4.layeronesoftware.com,1433"; Credential = "layeronedev" }
    @{ Name = "Convex-Test"; AppServer = "convex-test.layeronesoftware.com"; SqlServer = "convex-test.layeronesoftware.com,1433"; Credential = "convex" }
    @{ Name = "QA"; AppServer = "qa.layeronesoftware.com"; SqlServer = "qa.layeronesoftware.com,1433"; Credential = "layeronedev" }
)

$ProdEnvironments = @(
    @{ Name = "Client001"; AppServer = "client001.layeronesoftware.com"; SqlServer = "client001-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client001-UAT"; AppServer = "client001-uat.layeronesoftware.com"; SqlServer = "client001-sql-uat.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client003"; AppServer = "client003.layeronesoftware.com"; SqlServer = "client003.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client003-UAT"; AppServer = "client003-uat.layeronesoftware.com"; SqlServer = "client003-uat.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client007"; AppServer = "client007.layeronesoftware.com"; SqlServer = "client007-sql.layeronesoftware.com,1433"; Credential = "client007" }
    @{ Name = "Client007-UAT"; AppServer = "client007-uat.layeronesoftware.com"; SqlServer = "client007-sql-uat.layeronesoftware.com,1433"; Credential = "client007" }
    @{ Name = "Client009"; AppServer = "client009.layeronesoftware.com"; SqlServer = "client009-sql.layeronesoftware.com,1433"; Credential = "client009" }
    @{ Name = "Client011"; AppServer = "client011.layeronesoftware.com"; SqlServer = "client011-sql.layeronesoftware.com,1433"; Credential = "client011" }
    @{ Name = "Client012"; AppServer = "client012.layeronesoftware.com"; SqlServer = "client012-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client014"; AppServer = "client014.layeronesoftware.com"; SqlServer = "client014-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client014-UAT"; AppServer = "client014-uat.layeronesoftware.com"; SqlServer = "client014-sql-uat.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client018"; AppServer = "client018.layeronesoftware.com"; SqlServer = "sql2.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client020"; AppServer = "client020.layeronesoftware.com"; SqlServer = "sql4.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client021"; AppServer = "client021.layeronesoftware.com"; SqlServer = "client021-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client021-UAT"; AppServer = "client021-uat.layeronesoftware.com"; SqlServer = "client021-sql-uat.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client023"; AppServer = "client023.layeronesoftware.com"; SqlServer = "client023.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client024"; AppServer = "client024.layeronesoftware.com"; SqlServer = "client024-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client024-UAT"; AppServer = "client024-uat.layeronesoftware.com"; SqlServer = "client024-sql-uat.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client024-QA"; AppServer = "client024-qa.layeronesoftware.com"; SqlServer = "client024-sql-qa.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client025"; AppServer = "client025.layeronesoftware.com"; SqlServer = "client025-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client025-UAT"; AppServer = "client025-uat.layeronesoftware.com"; SqlServer = "client025-sql-uat.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client026"; AppServer = "client026.layeronesoftware.com"; SqlServer = "sql4.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client027"; AppServer = "client027.layeronesoftware.com"; SqlServer = "sql.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client028"; AppServer = "client028.layeronesoftware.com"; SqlServer = "sql2.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client029"; AppServer = "client029.layeronesoftware.com"; SqlServer = "sql5.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client030"; AppServer = "client030.layeronesoftware.com"; SqlServer = "client030-sql.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client031"; AppServer = "client031.layeronesoftware.com"; SqlServer = "client031-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client031-UAT"; AppServer = "client031-uat.layeronesoftware.com"; SqlServer = "client031-sql-uat.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client032"; AppServer = "client032.layeronesoftware.com"; SqlServer = "sql5.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client033"; AppServer = "client033.layeronesoftware.com"; SqlServer = "sql6.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client034"; AppServer = "client034.layeronesoftware.com"; SqlServer = "sql6.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client035"; AppServer = "client035.layeronesoftware.com"; SqlServer = "sql3.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client036"; AppServer = "client036.layeronesoftware.com"; SqlServer = "sql7.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client037"; AppServer = "client037.layeronesoftware.com"; SqlServer = "sql7.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client038"; AppServer = "client038.layeronesoftware.com"; SqlServer = "client038-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client039"; AppServer = "client039.layeronesoftware.com"; SqlServer = "sql8.layeronesoftware.com,1435"; Credential = "layeronecloud" }
    @{ Name = "Client040"; AppServer = "client040.layeronesoftware.com"; SqlServer = "client040-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client041"; AppServer = "client041.layeronesoftware.com"; SqlServer = "client041-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client042"; AppServer = "client042.layeronesoftware.com"; SqlServer = "client042-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }  
    @{ Name = "Client043"; AppServer = "client043.layeronesoftware.com"; SqlServer = "client043-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client044"; AppServer = "client044.layeronesoftware.com"; SqlServer = "client044-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client045"; AppServer = "client045.layeronesoftware.com"; SqlServer = "client045-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client046"; AppServer = "client046.layeronesoftware.com"; SqlServer = "client046-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client047"; AppServer = "client047.layeronesoftware.com"; SqlServer = "client047-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client048"; AppServer = "client048.layeronesoftware.com"; SqlServer = "client048-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client049"; AppServer = "client049.layeronesoftware.com"; SqlServer = "client049-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client050"; AppServer = "client050.layeronesoftware.com"; SqlServer = "client050-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client051"; AppServer = "client051.layeronesoftware.com"; SqlServer = "client051-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client052"; AppServer = "client052.layeronesoftware.com"; SqlServer = "client052-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client053"; AppServer = "client053.layeronesoftware.com"; SqlServer = "client053-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
    @{ Name = "Client054"; AppServer = "client054.layeronesoftware.com"; SqlServer = "client054-sql.layeronesoftware.com,1433"; Credential = "layeronecloud" }
)

foreach ($environment in $DevEnvironments)
{
    $env = $environment.Name
    $AppServerPath = "\\{0}\f$" -f $environment.AppServer
    $AppServerCMEntry = $environment.AppServer
    $SqlServerCMEntry = $environment.SqlServer -replace ",", ":"
    $SqlServerName = $environment.SqlServer
    $username = $CredentialsConstants[$environment.Credential].Username
    $password = $CredentialsConstants[$environment.Credential].Password

    CreateShortcut -ShortcutName $env -TargetPath $AppServerPath -OutputFolder $DevFolder

    CreateSSMSEntry -ServerGroup "Dev" -ServerName $env -ServerAddress $SqlServerName

    Add-Credential -Target $AppServerCMEntry -Username $username -Password $password
    Add-Credential -Target $SqlServerCMEntry -Username $username -Password $password
}

foreach ($environment in $ProdEnvironments)
{
    $env = $environment.Name
    $AppServerPath = "\\{0}" -f $environment.AppServer
    $AppServerCMEntry = $environment.AppServer
    $SqlServerCMEntry = $environment.SqlServer -replace ",", ":"
    $SqlServerName = $environment.SqlServer
    $username = $CredentialsConstants[$environment.Credential].Username
    $password = $CredentialsConstants[$environment.Credential].Password

    CreateShortcut -ShortcutName $env -TargetPath $AppServerPath -OutputFolder $ProdFolder

    CreateSSMSEntry -ServerGroup "Prod" -ServerName $env -ServerAddress $SqlServerName

    # to overwrite, call Add-Credential instead of Safe-Add-Credential
    Add-Credential -Target $AppServerCMEntry -Username $username -Password $password
    Add-Credential -Target $SqlServerCMEntry -Username $username -Password $password
}

Write-Host "Environment setup complete."

Pause