# PS1 setup script
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

if (-not (Get-Command "python" -ErrorAction SilentlyContinue)) {
    Write-Error "Python is not installed. Please install Python first."
    exit 1
}

# Check for Pandoc
$pandocSystem = Get-Command "pandoc" -ErrorAction SilentlyContinue
$pandocLocal = "$env:LOCALAPPDATA\Pandoc\pandoc.exe"

if (-not $pandocSystem -and -not (Test-Path $pandocLocal)) {
    Write-Host "Pandoc not found. Attempting to install via Winget..."
    winget install JohnMacFarlane.Pandoc --source winget
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Failed to install Pandoc via winget. Please install it manually from https://pandoc.org/installing.html"
    } else {
        Write-Host "Pandoc installed successfully."
    }
} else {
    Write-Host "Pandoc is available!"
}

# Virtual environment path
$VenvName = "latex_to_onenote_venv"
$VenvPath = Join-Path $env:USERPROFILE ".virtualenvs\$VenvName"

# Setup venv
if (-not (Test-Path $VenvPath)) {
    Write-Host "Creating virtual environment at $VenvPath..." -ForegroundColor Cyan
    python -m venv $VenvPath
} else {
    Write-Host "Virtual environment '$VenvName' already exists at $VenvPath." -ForegroundColor Green
}

# Install requirements
Write-Host "Installing/Updating Python dependencies..."
$pythonPath = Join-Path $VenvPath "Scripts\python.exe"

if (Test-Path $pythonPath) {
    & $pythonPath -m pip install --upgrade pip
    & $pythonPath -m pip install pyperclip python-docx pywin32 lxml
} else {
    Write-Error "Could not find python in venv. Please remove '$VenvPath' folder and run this script again."
}

Write-Host "`nSetup complete!"
Write-Host "To run the tool:"
Write-Host "1. & '$VenvPath\Scripts\Activate.ps1'"
Write-Host "2. python latex_to_onenote.py '<your_latex>'"

pause
