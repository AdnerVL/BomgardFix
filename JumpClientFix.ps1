# Define parameters at the script level
$Hostname = $args[0]

# If no hostname provided via argument, prompt for input
if (-not $Hostname) {
    $Hostname = Read-Host "Please enter the hostname"
}

# Ensure C:\Tools\Script directory exists
if (-Not (Test-Path -Path "C:\Tools\Script")) {
    New-Item -ItemType Directory -Path "C:\Tools\Script" | Out-Null
}

# Download PSExec if not exists
$psexecZipPath = "C:\Tools\Script\PSTools.zip"
$psexecExePath = "C:\Tools\Script\PsExec.exe"
$psexec64ExePath = "C:\Tools\Script\PsExec64.exe"

if (-Not (Test-Path $psexecExePath) -or -Not (Test-Path $psexec64ExePath)) {
    try {
        Write-Output "Step 1: Downloading PSExec..."
        Invoke-WebRequest -Uri "https://download.sysinternals.com/files/PSTools.zip" -OutFile $psexecZipPath

        Write-Output "Step 2: Extracting PSExec.exe and PSExec64.exe..."
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($psexecZipPath)
        foreach ($entry in $zipArchive.Entries) {
            if ($entry.FullName -eq "PsExec.exe" -or $entry.FullName -eq "PsExec64.exe") {
                $destinationPath = Join-Path "C:\Tools\Script" $entry.FullName
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destinationPath, $true)
            }
        }
        $zipArchive.Dispose()

        Write-Output "Step 3: Cleaning up zip file..."
        Remove-Item $psexecZipPath -Force
    }
    catch {
        Write-Error "Failed to download or extract PSExec: $_"
        Write-Output "Cleaning up by removing the Script folder..."
        Remove-Item "C:\Tools\Script" -Recurse -Force
        exit 1
    }
}

Write-Output "Step 4: Loading environment variables..."
# Load environment variables
$envFile = Join-Path $PSScriptRoot ".env"
$env:DOWNLOAD_URL = $null
$env:KEY_SECRET = $null

if (Test-Path $envFile) {
    Write-Output "Environment file found. Loading variables..."
    # Load environment variables from .env file
    Get-Content $envFile | ForEach-Object {
        if ($_ -match "^\s*([^#;].*?)\s*=\s*(.*?)\s*$") {
            $name, $value = $matches[1], $matches[2]
            ${env:$name} = $value
        }
    }
} else {
    Write-Output "No environment file found. Skipping variable loading..."
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

# Remote uninstallation and installation
try {
    # Prepare remote commands
    $remoteCommands = @(
        "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'Get-WmiObject Win32_Product | Where-Object { `$_.Name -like ''BeyondTrust Jump Client*'' } | ForEach-Object { Start-Process ''msiexec.exe'' -ArgumentList ''/x '', `$_.IdentifyingNumber, ''/qn /norestart'' -Wait }'`"",
        "if not exist C:\Tools\Script mkdir C:\Tools\Script",
        "powershell -Command `"Start-Process powershell -Verb RunAs -ArgumentList '-Command', 'Invoke-WebRequest -Uri ''$env:DOWNLOAD_URL'' -OutFile ''C:\Tools\Script\install.msi'''`"",
        "msiexec /i C:\Tools\Script\install.msi KEY_INFO=`"$env:KEY_SECRET`" /qn /norestart /l*v C:\Tools\Script\install.log"
    )

    Set-Location "C:\Tools\Script"  # Ensure we are in the directory where PsExec is located

    foreach ($command in $remoteCommands) {
        if ($command.Trim() -ne "") {
            Write-Output "Executing command: $command"
            $process = Start-Process -FilePath ".\PsExec.exe" -ArgumentList "\\$Hostname", "cmd", "/c", "`"$command`"" -Wait -PassThru -RedirectStandardOutput "C:\Tools\Script\psexec_output.txt" -RedirectStandardError "C:\Tools\Script\psexec_error.txt"
            if ($process.ExitCode -ne 0) {
                $errorContent = Get-Content "C:\Tools\Script\psexec_error.txt"
                $installLog = Get-Content "C:\Tools\Script\install.log" -ErrorAction SilentlyContinue
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
        Write-Error "Failed to clean up local PSExec files and Script folder: $_"
    }
}