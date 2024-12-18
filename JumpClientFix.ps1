# Define parameters at the script level instead of using param() block
$Hostname = $args[0]

# If no hostname provided via argument, prompt for input
if (-not $Hostname) {
    $Hostname = Read-Host "Please enter the hostname"
}

# Function to find PSExec in common locations
function Find-PSExec {
    $possibleLocations = @(
        $PSScriptRoot,  # Current script directory
        (Get-Location).Path,  # Current working directory
        "C:\Tools\Sysinternals",
        "C:\Program Files\Sysinternals",
        "C:\Program Files (x86)\Sysinternals",
        "$env:USERPROFILE\Downloads\Sysinternals"
    )

    foreach ($location in $possibleLocations) {
        # Try both PsExec.exe and PsExec64.exe
        $psexecPaths = @(
            (Join-Path $location "PsExec.exe"),
            (Join-Path $location "PsExec64.exe")
        )

        foreach ($psexecPath in $psexecPaths) {
            if (Test-Path $psexecPath) {
                return $psexecPath
            }
        }
    }

    return $null
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

# Ensure C:\Tools directory exists
if (-Not (Test-Path -Path "C:\Tools")) {
    New-Item -ItemType Directory -Path "C:\Tools" | Out-Null
}

# Function to get a standardized filename
function Get-StandardizedFileName {
    param([string]$OriginalFileName)
    
    # Remove any characters that might cause issues
    $sanitizedName = $OriginalFileName -replace '[^a-zA-Z0-9\.]',''
    
    # If the name is too long, truncate it
    if ($sanitizedName.Length -gt 50) {
        $sanitizedName = $sanitizedName.Substring(0, 50)
    }
    
    # Ensure it ends with .msi if it doesn't already
    if (-not $sanitizedName.ToLower().EndsWith('.msi')) {
        $sanitizedName += '.msi'
    }
    
    return $sanitizedName
}

# Find PSExec
$psexecPath = Find-PSExec

# Attempt to connect to hostname using PSExec
try {
    if (-not $psexecPath) {
        throw "PSExec not found. Please ensure PsExec.exe or PsExec64.exe is in the script directory or system PATH."
    }

    # Use full path to psexec
    $psexecCommand = "&`"$psexecPath`" \\$Hostname cmd /c `"echo Connection successful`""
    
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

# Download the file
try {
    $response = Invoke-WebRequest -Uri $env:DOWNLOAD_URL -UseBasicParsing
    
    # Get the filename from the response headers or URL
    $originalFileName = $response.Headers.'Content-Disposition' -replace '.*filename=', '' -replace '"',''
    if (-not $originalFileName) {
        $originalFileName = [System.IO.Path]::GetFileName($env:DOWNLOAD_URL)
    }
    
    # Standardize the filename
    $standardFileName = Get-StandardizedFileName -OriginalFileName $originalFileName
    
    # Save the file
    $filePath = Join-Path "C:\Tools" $standardFileName
    [System.IO.File]::WriteAllBytes($filePath, $response.Content)
    
    Write-Output "File downloaded to $filePath"
}
catch {
    Write-Error "Failed to download file: $_"
    exit 1
}

# Install the MSI
try {
    $installArgs = "/i `"$filePath`" KEY_INFO=`"$env:KEY_SECRET`" /qn /norestart"
    Start-Process msiexec.exe -ArgumentList $installArgs -Wait -PassThru
    
    Write-Output "Installation completed successfully"
}
catch {
    Write-Error "Installation failed: $_"
    exit 1
}