#requires -Version 5.1
$ErrorActionPreference = "Stop"

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = Split-Path -Parent $scriptDir
$configCandidates = @(
  (Join-Path $projectRoot "config\Syncro-TechSummary.config.json"),
  (Join-Path $scriptDir "Syncro-TechSummary.config.json"),
  (Join-Path $projectRoot "Syncro-TechSummary.config.json")
)
$configPath = $configCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($configPath)) { throw "Config not found. Checked: $($configCandidates -join ', ')" }
$config = Get-Content $configPath -Raw | ConvertFrom-Json

$sub = [string]$config.Subdomain
$key = [string]$config.ApiKey
if ([string]::IsNullOrWhiteSpace($sub)) { throw "Config.Subdomain is empty." }
if ([string]::IsNullOrWhiteSpace($key) -or $key -like "PUT_*" -or $key -like "PASTE_*") { throw "Config.ApiKey is not set." }

$endpoints = @(
  "/api/v1/timelogs?api_key=$key&per_page=5&page=1",
  "/api/v1/ticket_timers?api_key=$key&per_page=5&page=1",
  "/api/v1/ticket_timers?api_key=$key&page=1",
  "/api/v1/ticket_charges?api_key=$key&per_page=5&page=1",
  "/api/v1/charges?api_key=$key&per_page=5&page=1",
  "/api/v1/line_items?api_key=$key&per_page=5&page=1",
  "/api/v1/worklogs?api_key=$key&per_page=5&page=1",
  "/api/v1/ticket_worklogs?api_key=$key&per_page=5&page=1"
)

foreach ($ep in $endpoints) {
  $url = "https://$sub.syncromsp.com$ep"
  try {
    $resp = Invoke-RestMethod $url
    $keys = @($resp.PSObject.Properties.Name) -join ", "
    Write-Host "OK  $ep"
    Write-Host "    Top-level keys: $keys"
    foreach ($k in $resp.PSObject.Properties.Name) {
      $val = $resp.PSObject.Properties[$k].Value
      if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {
        try { Write-Host "    $k count: $($val.Count)" } catch {}
      }
    }
  } catch {
    $msg = $_.Exception.Message
    if ($msg -match "404") {
      Write-Host "404 $ep"
    } elseif ($msg -match "401|403") {
      Write-Host "DENIED $ep ($msg)"
    } else {
      Write-Host "ERR $ep ($msg)"
    }
  }
}
