# ReadAloud Setup Script v2
# This script downloads Python (if needed), the latest code, and runs the app

$ErrorActionPreference = "Stop"

# Configuration
$PythonVersion = "3.12.2"
$PythonEmbedUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"
$GitHubRepo = "akamrt/readaloud"
$GitHubBranch = "main"

# Paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonDir = Join-Path $ScriptDir "python"
$CodeDir = Join-Path $ScriptDir "readaloud"
$RequirementsFile = Join-Path $CodeDir "requirements.txt"
$MainScript = Join-Path $CodeDir "main.py"
$TempDir = Join-Path $ScriptDir "temp"

# Create directories
if (-not (Test-Path $PythonDir)) { New-Item -ItemType Directory -Path $PythonDir | Out-Null }
if (-not (Test-Path $TempDir)) { New-Item -ItemType Directory -Path $TempDir | Out-Null }

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ReadAloud Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check for system Python
Write-Host "Checking for Python..." -ForegroundColor Cyan
$pythonExe = $null
$useSystemPython = $false

# Try to find Python in PATH
try {
    $pythonPath = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonPath) {
        $pythonExe = $pythonPath.Source
        $useSystemPython = $true
        Write-Host "Found system Python: $pythonExe" -ForegroundColor Green
    }
} catch {}

# If not in PATH, check common locations
if (-not $pythonExe) {
    $commonPaths = @(
        "C:\Python3*\python.exe",
        "C:\Program Files\Python3*\python.exe",
        "C:\Program Files (x86)\Python3*\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
        "$env:APPDATA\Python\Python3*\python.exe"
    )
    
    foreach ($pattern in $commonPaths) {
        $found = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) {
            $pythonExe = $found.FullName
            $useSystemPython = $true
            Write-Host "Found Python at: $pythonExe" -ForegroundColor Green
            break
        }
    }
}

# If still not found, use embedded Python
if (-not $pythonExe) {
    Write-Host "Python not found. Using embedded Python..." -ForegroundColor Yellow
    
    $pythonExe = Join-Path $PythonDir "python.exe"
    if (-not (Test-Path $pythonExe)) {
        Write-Host "Downloading Python $PythonVersion embeddable..." -ForegroundColor Yellow
        
        $pythonZip = Join-Path $TempDir "python-embed.zip"
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $PythonEmbedUrl -OutFile $pythonZip -UseBasicParsing
        } catch {
            Write-Host "ERROR: Failed to download Python. Please check your internet connection." -ForegroundColor Red
            Write-Host "Error details: $_" -ForegroundColor Red
            exit 1
        }
        
        Write-Host "Extracting Python..." -ForegroundColor Yellow
        try {
            Expand-Archive -Path $pythonZip -DestinationPath $PythonDir -Force
            Remove-Item $pythonZip -Force
        } catch {
            Write-Host "ERROR: Failed to extract Python." -ForegroundColor Red
            Write-Host "Error details: $_" -ForegroundColor Red
            exit 1
        }
        
        # Enable site-packages in embedded Python
        $pthFile = Join-Path $PythonDir "python312._pth"
        if (Test-Path $pthFile) {
            $content = Get-Content $pthFile
            $content = $content -replace "#import site", "import site"
            Set-Content -Path $pthFile -Value $content
        }
        
        Write-Host "Python installed successfully." -ForegroundColor Green
    }
}

# Test Python installation
Write-Host "Testing Python..." -ForegroundColor Cyan
try {
    $version = & $pythonExe --version 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Python test failed" }
    Write-Host "Python version: $version" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python is not working correctly." -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
    exit 1
}

# Ensure pip is available
Write-Host "Checking pip..." -ForegroundColor Cyan
$pipExe = if ($useSystemPython) { "pip" } else { Join-Path $PythonDir "Scripts\pip.exe" }

if ($useSystemPython) {
    # Check if pip is available
    try {
        & $pythonExe -m pip --version 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "pip not found" }
    } catch {
        Write-Host "Installing pip..." -ForegroundColor Yellow
        $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
        $getPipScript = Join-Path $TempDir "get-pip.py"
        
        try {
            Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipScript -UseBasicParsing
            & $pythonExe $getPipScript --no-warn-script-location
            if ($LASTEXITCODE -ne 0) { throw "pip installation failed" }
        } catch {
            Write-Host "ERROR: Failed to install pip." -ForegroundColor Red
            Write-Host "Error details: $_" -ForegroundColor Red
            exit 1
        }
    }
} else {
    if (-not (Test-Path $pipExe)) {
        Write-Host "Installing pip..." -ForegroundColor Yellow
        $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
        $getPipScript = Join-Path $TempDir "get-pip.py"
        
        try {
            Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipScript -UseBasicParsing
            & $pythonExe $getPipScript --no-warn-script-location
            if ($LASTEXITCODE -ne 0) { throw "pip installation failed" }
        } catch {
            Write-Host "ERROR: Failed to install pip." -ForegroundColor Red
            Write-Host "Error details: $_" -ForegroundColor Red
            exit 1
        }
    }
}

# Download or update code
Write-Host ""
Write-Host "Checking for code updates..." -ForegroundColor Cyan

# Check if git is available
$gitAvailable = $false
try { git --version 2>&1 | Out-Null; $gitAvailable = $LASTEXITCODE -eq 0 } catch {}

if ($gitAvailable -and (Test-Path (Join-Path $CodeDir ".git"))) {
    # Update using git
    Write-Host "Updating code using git..." -ForegroundColor Yellow
    Set-Location $CodeDir
    git pull origin $GitHubBranch
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: git pull failed. Continuing with existing code." -ForegroundColor Yellow
    }
    Set-Location $ScriptDir
} else {
    # Download code as zip
    Write-Host "Downloading latest code..." -ForegroundColor Yellow
    
    $zipUrl = "https://github.com/$GitHubRepo/archive/refs/heads/$GitHubBranch.zip"
    $zipFile = Join-Path $TempDir "readaloud.zip"
    
    try {
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing
    } catch {
        Write-Host "ERROR: Failed to download code." -ForegroundColor Red
        Write-Host "Error details: $_" -ForegroundColor Red
        exit 1
    }
    
    # Extract code
    if (Test-Path $CodeDir) {
        Remove-Item -Path $CodeDir -Recurse -Force
    }
    
    Write-Host "Extracting code..." -ForegroundColor Yellow
    try {
        Expand-Archive -Path $zipFile -DestinationPath $TempDir -Force
        $extractedDir = Join-Path $TempDir "readaloud-$GitHubBranch"
        Move-Item -Path $extractedDir -Destination $CodeDir
        Remove-Item $zipFile -Force
    } catch {
        Write-Host "ERROR: Failed to extract code." -ForegroundColor Red
        Write-Host "Error details: $_" -ForegroundColor Red
        exit 1
    }
}

# Install dependencies
Write-Host ""
Write-Host "Installing dependencies..." -ForegroundColor Cyan
if (Test-Path $RequirementsFile) {
    & $pythonExe -m pip install -r $RequirementsFile --no-warn-script-location
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Some dependencies may have failed to install." -ForegroundColor Yellow
    } else {
        Write-Host "Dependencies installed successfully." -ForegroundColor Green
    }
} else {
    Write-Host "Warning: requirements.txt not found." -ForegroundColor Yellow
}

# Check for optional AI OCR (auto-install if user agrees)
Write-Host ""
Write-Host "Checking optional AI OCR..." -ForegroundColor Cyan
$aiRequirements = Join-Path $CodeDir "requirements-ai.txt"
if (Test-Path $aiRequirements) {
    Write-Host "AI OCR provides better accuracy for complex text layouts." -ForegroundColor Yellow
    $installAI = Read-Host "Install AI OCR? (y/N)"
    if ($installAI -eq "y" -or $installAI -eq "Y") {
        Write-Host "Installing AI OCR dependencies..." -ForegroundColor Yellow
        & $pythonExe -m pip install -r $aiRequirements --no-warn-script-location
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Warning: AI OCR installation failed. Continuing without it." -ForegroundColor Yellow
        } else {
            Write-Host "AI OCR installed successfully." -ForegroundColor Green
        }
    } else {
        Write-Host "Skipping AI OCR installation." -ForegroundColor Yellow
    }
}

# Create desktop shortcut (optional)
Write-Host ""
$createShortcut = Read-Host "Create desktop shortcut? (y/N)"
if ($createShortcut -eq "y" -or $createShortcut -eq "Y") {
    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\ReadAloud.lnk")
        $Shortcut.TargetPath = $pythonExe
        $Shortcut.Arguments = "`"$MainScript`""
        $Shortcut.WorkingDirectory = $CodeDir
        $Shortcut.Description = "ReadAloud - Text-to-Speech Tool"
        $Shortcut.Save()
        Write-Host "Desktop shortcut created." -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not create shortcut." -ForegroundColor Yellow
    }
}

# Run the application
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Starting ReadAloud..." -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $MainScript) {
    & $pythonExe $MainScript
} else {
    Write-Host "ERROR: main.py not found at $MainScript" -ForegroundColor Red
    exit 1
}