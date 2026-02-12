#requires -Version 5.1
$ErrorActionPreference = "Stop"

$sub = "bigfootnetworks"
$key = "REDACTED_SYNCRO_API_KEY"

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
