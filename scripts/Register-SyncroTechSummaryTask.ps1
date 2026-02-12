#requires -Version 5.1
$ErrorActionPreference = "Stop"

param(
    [string]$TaskName = "SyncroTechSummary",
    [int]$IntervalMinutes = 30,
    [switch]$UseCurrentUser,
    [string]$RunAsUser
)

if ($IntervalMinutes -lt 1) {
    throw "IntervalMinutes must be >= 1."
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$runner = Join-Path $repoRoot "Run-SyncroTechSummary.ps1"
if (!(Test-Path $runner)) {
    throw "Run script not found: $runner"
}

$runnerEscaped = $runner.Replace('"', '\"')
$tr = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$runnerEscaped`""

$args = @(
    "/Create",
    "/TN", $TaskName,
    "/SC", "MINUTE",
    "/MO", ([string]$IntervalMinutes),
    "/TR", $tr,
    "/F"
)

if ($UseCurrentUser) {
    # No /RU means "current user context".
} elseif (-not [string]::IsNullOrWhiteSpace($RunAsUser)) {
    $args += @("/RU", $RunAsUser, "/RP", "*")
}

Write-Host "Registering task '$TaskName'..."
& schtasks.exe @args
if ($LASTEXITCODE -ne 0) {
    throw "schtasks failed with exit code $LASTEXITCODE"
}

Write-Host "Task registered."
Write-Host "Run now (optional): schtasks /Run /TN `"$TaskName`""
