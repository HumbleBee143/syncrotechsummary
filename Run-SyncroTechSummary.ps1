#requires -Version 5.1
$ErrorActionPreference = "Stop"

$rootDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$mainScript = Join-Path $rootDir "scripts\Syncro-TechSummary.ps1"
if (!(Test-Path $mainScript)) { throw "Main script not found: $mainScript" }

& $mainScript
