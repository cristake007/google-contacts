[CmdletBinding()]
param(
    [string]$VenvDirectory = ".venv",

    [ValidateSet("None", "Fresh", "Resume", "RetryFailed")]
    [string]$RunMode = "None",

    [ValidateRange(0, 2147483647)]
    [int]$Limit = 20
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repositoryRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPath = Join-Path $repositoryRoot $VenvDirectory
$venvPython = Join-Path $venvPath "Scripts\python.exe"
$requirementsPath = Join-Path $repositoryRoot "requirements.txt"
$activateScript = Join-Path $venvPath "Scripts\Activate.ps1"
$runnerScript = Join-Path $repositoryRoot "run_contacts.ps1"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Creating virtual environment at $venvPath"

    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $venvPath
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv $venvPath
    }
    else {
        throw "Python 3.10 or newer was not found. Install Python, then rerun this script."
    }

    if ($LASTEXITCODE -ne 0) {
        throw "Virtual environment creation failed with exit code $LASTEXITCODE."
    }
}
else {
    Write-Host "Using existing virtual environment at $venvPath"
}

$pythonVersion = & $venvPython -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if ($LASTEXITCODE -ne 0) {
    throw "Unable to run the virtual environment's Python interpreter."
}

$versionParts = $pythonVersion.Trim().Split(".")
if ([int]$versionParts[0] -lt 3 -or ([int]$versionParts[0] -eq 3 -and [int]$versionParts[1] -lt 10)) {
    throw "Python 3.10 or newer is required; the virtual environment uses Python $pythonVersion."
}

Write-Host "Upgrading pip"
& $venvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    throw "pip upgrade failed with exit code $LASTEXITCODE."
}

Write-Host "Installing Python packages from requirements.txt"
& $venvPython -m pip install --requirement $requirementsPath
if ($LASTEXITCODE -ne 0) {
    throw "Package installation failed with exit code $LASTEXITCODE."
}

Write-Host "Installing Playwright Chromium"
& $venvPython -m playwright install chromium
if ($LASTEXITCODE -ne 0) {
    throw "Chromium installation failed with exit code $LASTEXITCODE."
}

Write-Host ""
Write-Host "Activating the virtual environment"
. $activateScript

Write-Host "Setup complete."
Write-Host "The virtual environment is active. To reactivate it in a later shell:"
Write-Host "  .\$VenvDirectory\Scripts\Activate.ps1"
Write-Host ""
Write-Host "Start fresh with:"
Write-Host "  .\run_contacts.ps1 -Mode Fresh -Limit $Limit"
Write-Host "Resume with:"
Write-Host "  .\run_contacts.ps1 -Mode Resume -Limit $Limit"
Write-Host "Resume and retry failures/review rows with:"
Write-Host "  .\run_contacts.ps1 -Mode RetryFailed -Limit $Limit"

if ($RunMode -ne "None") {
    Write-Host ""
    & $runnerScript -Mode $RunMode -Limit $Limit -VenvDirectory $VenvDirectory
}
