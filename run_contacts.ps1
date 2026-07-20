[CmdletBinding()]
param(
    [ValidateSet("Fresh", "Resume", "RetryFailed")]
    [string]$Mode = "Fresh",

    [string]$InputWorkbook = "all_cleaned.xlsx",

    [string]$VenvDirectory = ".venv",

    [ValidateRange(0, 2147483647)]
    [int]$Limit = 20,

    [string]$OnlyCui = "",

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AdditionalArguments = @()
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repositoryRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPath = Join-Path $repositoryRoot $VenvDirectory
$venvPython = Join-Path $venvPath "Scripts\python.exe"
$finderScript = Join-Path $repositoryRoot "google_contact_finder_v4.py"

if (-not (Test-Path -LiteralPath $venvPython)) {
    throw "Virtual environment not found. Run .\bootstrap.ps1 first."
}

if ([System.IO.Path]::IsPathRooted($InputWorkbook)) {
    $inputPath = $InputWorkbook
}
else {
    $inputPath = Join-Path $repositoryRoot $InputWorkbook
}

if (-not (Test-Path -LiteralPath $inputPath)) {
    throw "Input workbook not found: $inputPath"
}

$finderArguments = @($finderScript, $inputPath, "--limit", $Limit)

switch ($Mode) {
    "Resume" {
        $finderArguments += "--resume"
    }
    "RetryFailed" {
        $finderArguments += @("--resume", "--retry-failed")
    }
}

if ($OnlyCui) {
    $finderArguments += @("--only-cui", $OnlyCui)
}

if ($AdditionalArguments.Count -gt 0) {
    $finderArguments += $AdditionalArguments
}

Write-Host "Starting Google Contact Finder in $Mode mode"
& $venvPython @finderArguments
if ($LASTEXITCODE -ne 0) {
    throw "Google Contact Finder exited with code $LASTEXITCODE."
}
