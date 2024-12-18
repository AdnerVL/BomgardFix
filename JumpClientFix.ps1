# Define parameters at the script level
$Hostname = $args[0]

# If no hostname provided via argument, prompt for input
if (-not $Hostname) {
    $Hostname = Read-Host "Please enter the hostname"
}

# Ensure C:\Tools directory exists
if (-Not (Test-Path -Path "C:\Tools")) {
    New-Item -ItemType Directory -Path "C:\Tools" | Out-Null
}

# Download PSExec if not exists
$psexecZipPath = "C:\Tools\PSTools.zip"
$psexecExePath = "C:\Tools\PsExec.exe"

if (-Not (Test-Path $psexecExePath)) {
    try {
        Write-Output "Downloading PSExec..."
        Invoke-WebRequest -Uri "https://download.sysinternals.com/files/PSTools.zip" -OutFile $psexecZipPath

        # Extract PSExec
        Expand-Archive -Path $psexecZipPath -DestinationPath "C:\Tools" -Force

        # Clean up zip file
        Remove-Item $psexecZipPath -Force
    }
    catch {
        Write-Error "Failed to download PSExec: $_"
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
    Set-Location "C:\Tools"
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

# Remote uninstallation and installation
try {
    # Prepare remote commands
    $remoteCommands = @(
        "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'Get-WmiObject Win32_Product | Where-Object { `$_.Name -like ''BeyondTrust Jump Client*'' } | ForEach-Object { Start-Process ''msiexec.exe'' -ArgumentList ''/x '', `$_.IdentifyingNumber, ''/qn /norestart'' -Wait }'`"",
        "if not exist C:\Tools mkdir C:\Tools",
        "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'Invoke-WebRequest -Uri ''$env:DOWNLOAD_URL'' -OutFile ''C:\Tools\install.msi'''`"",
        "msiexec /i C:\Tools\install.msi KEY_INFO=`"$env:KEY_SECRET`" /qn /norestart /l*v C:\Tools\install.log"
    )

    Set-Location "C:\Tools"  # Ensure we are in the directory where PsExec is located

    foreach ($command in $remoteCommands) {
        if ($command.Trim() -ne "") {
            Write-Output "Executing command: $command"
            $process = Start-Process -FilePath ".\PsExec.exe" -ArgumentList "\\$Hostname", "cmd", "/c", "`"$command`"" -Wait -PassThru -RedirectStandardOutput "C:\Tools\psexec_output.txt" -RedirectStandardError "C:\Tools\psexec_error.txt"
            if ($process.ExitCode -ne 0) {
                $errorContent = Get-Content "C:\Tools\psexec_error.txt"
                $installLog = Get-Content "C:\Tools\install.log" -ErrorAction SilentlyContinue
                throw "Remote command failed with exit code $($process.ExitCode). Error details:`n$errorContent`nInstallation log:`n$installLog"
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
    # Cleanup: Remove local PSExec files
    try {
        # Remove PSExec executables
        Remove-Item "C:\Tools\PsExec.exe" -Force
        Remove-Item "C:\Tools\PsExec64.exe" -ErrorAction SilentlyContinue
        Remove-Item "C:\Tools\psexec_output.txt" -ErrorAction SilentlyContinue
        Remove-Item "C:\Tools\psexec_error.txt" -ErrorAction SilentlyContinue
        
        Write-Output "Cleaned up local PSExec files"
    }
    catch {
        Write-Error "Failed to clean up local PSExec files: $_"
    }
}