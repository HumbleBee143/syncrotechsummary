$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$path = Join-Path $scriptDir "Syncro-TechSummary.ps1"
$open = @()
. $path
# not running full script. re-read current open tickets from API only
