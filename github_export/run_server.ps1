$ErrorActionPreference = "Stop"

Write-Host "Starting Document Assistant..." -ForegroundColor Green

# Check if venv exists
if (-not (Test-Path ".venv")) {
    Write-Host "Creating virtual environment..."
    python -m venv .venv
}

# Activate venv
.\.venv\Scripts\Activate.ps1

# Install requirements
Write-Host "Installing dependencies..."
pip install -r requirements.txt

# Run server
Write-Host "Starting Flask Server..."
$env:FLASK_APP = "work_assistant/txtapp.py"
$env:FLASK_DEBUG = "1"
python work_assistant/txtapp.py
