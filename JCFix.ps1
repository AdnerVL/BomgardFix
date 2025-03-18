param (
    [Parameter(Mandatory=$false)]
    [string]$hostname
)

# Logging function
function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logPath = "C:\Tools\Script\script.log"
    if (-not (Test-Path "C:\Tools\Script")) {
        New-Item -ItemType Directory -Path "C:\Tools\Script" -Force -ErrorAction Stop | Out-Null
    }
    "$timestamp - $Message" | Out-File -FilePath $logPath -Append -ErrorAction SilentlyContinue
}

# Validate hostname (allow single-label names)
function Test-Hostname {
    param ([string]$Name)
    $ipPattern = "^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    $fqdnPattern = "^(?:(?!-)[A-Za-z0-9-]{1,63}(?<!-)\.)+[A-Za-z]{2,}$"
    $singleLabelPattern = "^[A-Za-z0-9-]{1,63}$"
    return ($Name -match $ipPattern -or $Name -match $fqdnPattern -or $Name -match $singleLabelPattern)
}

# Download PSExec
function Download-PsExec {
    param ([string]$ZipPath, [string]$ExePath)
    if (-not (Test-Path $ExePath)) {
        Write-Log "Starting PSExec download..."
        Write-Progress -Activity "Downloading PSExec" -Status "In Progress" -PercentComplete 0
        Invoke-WebRequest -Uri "https://download.sysinternals.com/files/PSTools.zip" -OutFile $ZipPath -TimeoutSec 300 -ErrorAction Stop
        Write-Progress -Activity "Downloading PSExec" -Status "Extracting" -PercentComplete 50
        Expand-Archive -Path $ZipPath -DestinationPath "C:\Tools\Script" -Force -ErrorAction Stop
        Get-ChildItem "C:\Tools\Script" -Exclude "PsExec.exe" | Remove-Item -Force -ErrorAction Stop
        Write-Progress -Activity "Downloading PSExec" -Completed
        Write-Log "PSExec downloaded and extracted"
    }
}

# Execute remote command with timeout and progress
function Invoke-RemoteCommand {
    param ([string]$TargetHost, [string]$Command, [int]$Timeout = 300, [int]$ProgressPercent)
    Write-Log "Executing on ${TargetHost}: $Command"
    Write-Progress -Activity "Remote Operation on $TargetHost" -Status "Running Command" -PercentComplete $ProgressPercent
    $process = Start-Process -FilePath ".\PsExec.exe" -ArgumentList "-accepteula", "-s", "\\$TargetHost", "cmd", "/c", "`"$Command`"" -Wait -PassThru -NoNewWindow -RedirectStandardOutput "C:\Tools\Script\psexec_output.txt" -RedirectStandardError "C:\Tools\Script\psexec_error.txt" -ErrorAction Stop
    $output = Get-Content "C:\Tools\Script\psexec_output.txt" -ErrorAction SilentlyContinue
    $errorContent = Get-Content "C:\Tools\Script\psexec_error.txt" -ErrorAction SilentlyContinue
    if ($process.ExitCode -ne 0) {
        Write-Log "Command failed with exit code $($process.ExitCode): $errorContent"
        return $process.ExitCode, $errorContent
    }
    Write-Log "Command completed successfully: $output"
    return 0, $output
}

# Main script
try {
    Write-Log "Script execution started"

    # Hostname validation
    if ([string]::IsNullOrWhiteSpace($hostname)) {
        $hostname = Read-Host "Enter hostname"
    }
    if (-not (Test-Hostname $hostname)) {
        Write-Log "Invalid hostname: $hostname"
        Write-Error "Invalid hostname format: $hostname"
        exit 2
    }
    Write-Log "Using hostname: $hostname"

    # PSExec setup
    $psexecZipPath = "C:\Tools\Script\PSTools.zip"
    $psexecExePath = "C:\Tools\Script\PsExec.exe"
    Download-PsExec -ZipPath $psexecZipPath -ExePath $psexecExePath

    # Load environment variables
    $envFile = Join-Path $PSScriptRoot ".env"
    $env:DOWNLOAD_URL = $null
    $secureKey = $null
    if (Test-Path $envFile) {
        foreach ($line in (Get-Content $envFile)) {
            $line = $line.Trim()
            if ($line -and $line -notlike '#*') {
                $key, $value = $line -split '=', 2
                $key = $key.Trim()
                $value = $value.Trim().Trim('"''')
                switch ($key) {
                    "DOWNLOAD_URL" { $env:DOWNLOAD_URL = $value }
                    "KEY_SECRET" { $secureKey = ConvertTo-SecureString $value -AsPlainText -Force }
                }
            }
        }
    }
    if (-not $env:DOWNLOAD_URL -or -not $secureKey) {
        Write-Log "Missing DOWNLOAD_URL or KEY_SECRET"
        Write-Error "Missing required environment variables"
        exit 3
    }
    Write-Log "Environment variables loaded"

    # Test connection
    Set-Location "C:\Tools\Script"
    $exitCode, $errorContent = Invoke-RemoteCommand -TargetHost $hostname -Command "echo Connection successful" -ProgressPercent 25
    if ($exitCode -ne 0) { throw "Connection failed: $errorContent" }
    Write-Output "Connected to $hostname"

    # Check installation process
    $checkCmd = 'powershell.exe -Command "if (Get-Process msiexec -ErrorAction SilentlyContinue) { ''msiexec running'' } else { ''no msiexec'' }"'
    $exitCode, $processList = Invoke-RemoteCommand -TargetHost $hostname -Command $checkCmd -ProgressPercent 50
    if ($exitCode -ne 0) { throw "Process check failed: $processList" }
    if ($processList -match "msiexec running") {
        Write-Log "Installation process detected on $hostname"
        Write-Output "Installation process found"
        $stop = Read-Host "Stop installation? (y/n)"
        if ($stop -eq 'y') {
            $exitCode, $errorContent = Invoke-RemoteCommand -TargetHost $hostname -Command "powershell -Command 'Stop-Process -Name msiexec -Force'" -ProgressPercent 60
            if ($exitCode -ne 0) { throw "Failed to stop installation: $errorContent" }
            Write-Output "Installation stopped"
        }
    } else {
        Write-Log "No installation process on $hostname"
        Write-Output "No installation process found"
    }

    # Remote uninstall/install
    $remoteCommands = @(
        @{Cmd = "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'Get-WmiObject Win32_Product | Where-Object { `$_.Name -like ''BeyondTrust Jump Client*'' } | ForEach-Object { Start-Process ''msiexec.exe'' -ArgumentList ''/x '', `$_.IdentifyingNumber, ''/qn /norestart'' -Wait }'`""; Percent = 75},
        @{Cmd = "if not exist C:\Tools\Script mkdir C:\Tools\Script"; Percent = 80},
        @{Cmd = "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'try { Invoke-WebRequest -Uri ''$($env:DOWNLOAD_URL)'' -OutFile ''C:\Tools\Script\install.msi'' -TimeoutSec 600 -ErrorAction Stop; Write-Output ''Download succeeded'' } catch { Write-Output ''Download failed: '' + `$_.Exception.Message }'`""; Percent = 85},
        @{Cmd = "powershell -Command 'if (Test-Path ''C:\Tools\Script\install.msi'') { ''MSI exists'' } else { ''MSI missing'' }'"; Percent = 87},
        @{Cmd = "msiexec /i C:\Tools\Script\install.msi KEY_INFO=`"$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)))`" /qn /norestart /l*v C:\Tools\Script\install.log"; Percent = 90}
    )
    $maxRetries = 3
    foreach ($cmd in $remoteCommands) {
        for ($retryCount = 0; $retryCount -lt $maxRetries; $retryCount++) {
            $exitCode, $output = Invoke-RemoteCommand -TargetHost $hostname -Command $cmd.Cmd -Timeout 600 -ProgressPercent $cmd.Percent
            if ($exitCode -eq 0) {
                if ($cmd.Percent -eq 85 -and $output -notmatch "Download succeeded") {
                    throw "MSI download failed on remote machine: $output"
                }
                if ($cmd.Percent -eq 87 -and $output -match "MSI missing") {
                    throw "MSI file not found on remote machine after download attempt"
                }
                break
            }
            if ($exitCode -eq 1618) {
                Write-Log "Retry $retryCount/$maxRetries due to MSI conflict (1618)"
                Start-Sleep -Seconds 20
            } elseif ($exitCode -eq 1619) {
                Write-Log "MSI file issue (1619): $output"
                throw "Installation failed: MSI file not found or invalid ($output)"
            } else {
                Write-Log "Command failed with exit code ${exitCode}: $output"
                throw "Command failed after $maxRetries retries: $output"
            }
            if ($retryCount -eq $maxRetries - 1) {
                throw "Max retries exceeded: $output"
            }
        }
    }
    Write-Log "Remote uninstall/install completed"
    Write-Output "Uninstall/install completed"
    Write-Progress -Activity "Remote Operation on $hostname" -Completed
} catch {
    Write-Log "Script failed: $($_.Exception.Message)"
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
} finally {
    $filesToRemove = @("PsExec.exe", "PSTools.zip", "psexec_output.txt", "psexec_error.txt", "script.log", "install.msi", "install.log")
    $retry = 0
    while ((Test-Path "C:\Tools\Script") -and $retry -lt 3) {
        foreach ($file in $filesToRemove) {
            Remove-Item "C:\Tools\Script\$file" -Force -ErrorAction SilentlyContinue
        }
        Remove-Item "C:\Tools\Script" -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path "C:\Tools\Script") {
            Write-Log "Cleanup retry $retry/3 failed"
            Start-Sleep -Seconds 10
            $retry++
        } else {
            Write-Log "Cleanup completed"
            Write-Output "Cleanup completed"
            break
        }
    }
    if (Test-Path "C:\Tools\Script") {
        Write-Log "Final cleanup failed, files may be in use"
        Write-Output "Cleanup incomplete, manual removal may be needed"
    }
    exit 0
}