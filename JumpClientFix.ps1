# Define parameters at the script level
param(
    [Parameter(Mandatory=$false)]
    [string]$hostname
)

# If hostname is not provided, prompt the user
if ([string]::IsNullOrWhiteSpace($hostname)) {
    $hostname = Read-Host "Please enter the hostname"
}

# Validate that hostname is not empty after prompting
if ([string]::IsNullOrWhiteSpace($hostname)) {
    Write-Error "Hostname cannot be empty. Script cannot continue."
    exit
}

# Ensure C:\Tools\Script directory exists
if (-Not (Test-Path -Path "C:\Tools\Script")) {
    New-Item -ItemType Directory -Path "C:\Tools\Script" | Out-Null
}

# Download PSExec if not exists
$psexecZipPath = "C:\Tools\Script\PSTools.zip"
$psexecExePath = "C:\Tools\Script\PsExec.exe"

if (-Not (Test-Path $psexecExePath)) {
    try {
        Write-Output "Step 1: Downloading PSExec..."
        Invoke-WebRequest -Uri "https://download.sysinternals.com/files/PSTools.zip" -OutFile $psexecZipPath

        Write-Output "Step 2: Extracting PSExec.exe..."
        # Use Expand-Archive if ZipFile type is not available
        Expand-Archive -Path $psexecZipPath -DestinationPath "C:\Tools\Script" -Force

        # Remove all files except psexec.exe
        Write-Output "Step 3: Cleaning up zip file..."
        Get-ChildItem "C:\Tools\Script" -Exclude "PsExec.exe" | Remove-Item -Force

        #Remove-Item $psexecZipPath -Force
    }
    catch {
        Write-Error "Failed to download or extract PSExec: $_"
        Write-Output "Cleaning up by removing the Script folder..."
        # Attempt to close any open handles or wait before removal
        Start-Sleep -Seconds 10
        try {
            Get-Process | Where-Object { $_.Path -like "C:\Tools\Script*" } | Stop-Process -Force -ErrorAction SilentlyContinue
            Remove-Item "C:\Tools\Script" -Recurse -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Error "Failed to remove Script folder: $_"
        }
        exit 1
    }
}

# Load environment variables
$envFile = Join-Path $PSScriptRoot ".env"
$env:DOWNLOAD_URL = $null
$env:KEY_SECRET = $null

if (Test-Path $envFile) {
    $envContent = Get-Content $envFile
    foreach ($line in $envContent) {
        $line = $line.Trim()
        if ($line -and $line -notlike '#*') {
            $key, $value = $line -split '=', 2
            $key = $key.Trim()
            $value = $value.Trim().Trim('"''')
            
            switch ($key) {
                "DOWNLOAD_URL" { $env:DOWNLOAD_URL = $value }
                "KEY_SECRET" { $env:KEY_SECRET = $value }
            }
        }
    }
}

# Validate required environment variables
if (-not $env:DOWNLOAD_URL -or -not $env:KEY_SECRET) {
    Write-Error "Missing DOWNLOAD_URL or KEY_SECRET in .env file"
    exit 1
}

# Attempt to connect to hostname using PSExec
try {
    Set-Location "C:\Tools\Script"
    $psexecCommand = ".\PsExec.exe \\$Hostname cmd /c `"echo Connection successful`""
    
    # Capture and display output
    $result = Invoke-Expression $psexecCommand 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        throw "PSExec connection failed: $result"
    }
    
    Write-Output "Successfully connected to $Hostname"
}
catch {
    Write-Error "Could not establish connection to $Hostname"
    Write-Error $_.Exception.Message
    exit 1
}

# Check for installation in progress
#$checkCommand = "powershell.exe -Command `"Get-Process msiexec`""
$checkCommand = 'cmd /c powershell.exe -Command "Get-Process msiexec"'
$processList = & $psexecExePath -accepteula -s \\$hostname $checkCommand

if ($processList) {
    Write-Output "Installation process found on ${hostname}:"
    Write-Output $processList

    $stop = Read-Host "Do you want to stop the installation process? (y/n)"
    if ($stop -eq 'y') {
        $stopCommand = "powershell -Command `"Stop-Process -Name msiexec -Force`""
        & $psexecExePath -accepteula -s \\$hostname $stopCommand
        Write-Output "Installation process stopped."
    } else {
        Write-Output "Installation process not stopped."
    }
} else {
    Write-Output "No installation process found on $hostname."
}

# Remote uninstallation and installation
try {
    # Prepare remote commands
    $remoteCommands = @(
        "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'Get-WmiObject Win32_Product | Where-Object { `$_.Name -like ''BeyondTrust Jump Client*'' } | ForEach-Object { Start-Process ''msiexec.exe'' -ArgumentList ''/x '', `$_.IdentifyingNumber, ''/qn /norestart'' -Wait }'`"",
        "if not exist C:\Tools\Script mkdir C:\Tools\Script",
        "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'Invoke-WebRequest -Uri ''$($env:DOWNLOAD_URL)'' -OutFile ''C:\Tools\Script\install.msi'''`"",
        "msiexec /i C:\Tools\Script\install.msi KEY_INFO=`"$($env:KEY_SECRET)`" /qn /norestart /l*v C:\Tools\Script\install.log"
    )

    Set-Location "C:\Tools\Script"  # Ensure we are in the directory where PsExec is located

    foreach ($command in $remoteCommands) {
        if ($command.Trim() -ne "") {
            Write-Output "Executing command: $command"
            $retryCount = 0
            $maxRetries = 3
            $waitSeconds = 20
            
            do {
                $process = Start-Process -FilePath ".\PsExec.exe" -ArgumentList "\\$Hostname", "cmd", "/c", "`"$command`"" -Wait -PassThru -RedirectStandardOutput "C:\Tools\Script\psexec_output.txt" -RedirectStandardError "C:\Tools\Script\psexec_error.txt"
                
                if ($process.ExitCode -eq 1618) {
                    Write-Output "Another installation is in progress. Waiting and will retry..."
                    Start-Sleep -Seconds $waitSeconds
                    $retryCount++
                } else {
                    break
                }
            } while ($retryCount -lt $maxRetries)
            
            if ($process.ExitCode -ne 0) {
                $errorContent = Get-Content "C:\Tools\Script\psexec_error.txt"
                $installLog = Get-Content "C:\Tools\Script\install.log" -ErrorAction SilentlyContinue
                Write-Output "Installation log content:"
                Write-Output $installLog
                throw "Remote command failed with exit code $($process.ExitCode). Error details:`n$errorContent"
            }
        }
    }
    Write-Output "Remote uninstallation and installation completed successfully"
}
catch {
    Write-Error "Failed to uninstall/install on remote host: $_"
    exit 1
}
finally {
    # Cleanup: Remove local PSExec files and Script folder
    try {
        # Remove PSExec executables
        Remove-Item "C:\Tools\Script\PsExec.exe" -Force
        Remove-Item "C:\Tools\Script\PsExec64.exe" -ErrorAction SilentlyContinue
        Remove-Item "C:\Tools\Script\psexec_output.txt" -ErrorAction SilentlyContinue
        Remove-Item "C:\Tools\Script\psexec_error.txt" -ErrorAction SilentlyContinue
        
        # Remove Script folder but keep Tools folder
        Remove-Item "C:\Tools\Script" -Recurse -Force
        Write-Output "Cleaned up local PSExec files and Script folder"
    } 
    catch {
        Write-Output "Failed to remove Script folder, possibly in use. Waiting..."
        Start-Sleep -Seconds 30
        Remove-Item "C:\Tools\Script" -Recurse -Force -ErrorAction SilentlyContinue
    }
}