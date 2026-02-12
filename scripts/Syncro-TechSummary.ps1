#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -----------------------------
# Load config
# -----------------------------
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = Split-Path -Parent $scriptDir
$configCandidates = @(
    (Join-Path $projectRoot "config\Syncro-TechSummary.config.local.json"),
    (Join-Path $projectRoot "config\Syncro-TechSummary.config.json"),
    (Join-Path $scriptDir "Syncro-TechSummary.config.json"),
    (Join-Path $projectRoot "Syncro-TechSummary.config.json")
)
$configPath = $configCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($configPath)) {
    throw "Config not found. Checked: $($configCandidates -join ', ')"
}
if (!(Test-Path $configPath)) { throw "Config not found: $configPath" }
$config = Get-Content $configPath -Raw | ConvertFrom-Json

$subdomain = [string]$config.Subdomain
$apiKeyFromEnv = [string]$env:SYNCRO_API_KEY
$apiKey    = if ([string]::IsNullOrWhiteSpace($apiKeyFromEnv)) { [string]$config.ApiKey } else { $apiKeyFromEnv }
$daysBack  = if ($config.PSObject.Properties.Name -contains 'DaysBack') { [int]$config.DaysBack } else { 1 }

$windowMode = "RollingDays"
if ($config.PSObject.Properties.Name -contains 'Window' -and $config.Window) {
    if ($config.Window.PSObject.Properties.Name -contains 'Mode') { $windowMode = [string]$config.Window.Mode }
}

$windowEndNow = $true
$windowEndAtUtcMidnight = $false
if ($config.PSObject.Properties.Name -contains 'Window' -and $config.Window) {
    if ($config.Window.PSObject.Properties.Name -contains 'EndAtNow') { $windowEndNow = [bool]$config.Window.EndAtNow }
    if ($config.Window.PSObject.Properties.Name -contains 'EndAtUtcMidnight') { $windowEndAtUtcMidnight = [bool]$config.Window.EndAtUtcMidnight }
}

$reportPath =
if (($config.PSObject.Properties.Name -contains 'Output') -and $config.Output -and ($config.Output.PSObject.Properties.Name -contains 'ReportPath')) {
    [string]$config.Output.ReportPath
} else {
    Join-Path $projectRoot "output\LatestReport.txt"
}
if ([string]::IsNullOrWhiteSpace($reportPath)) {
    $reportPath = Join-Path $projectRoot "output\LatestReport.txt"
} elseif (-not [System.IO.Path]::IsPathRooted($reportPath)) {
    $reportPath = Join-Path $projectRoot $reportPath
}

$topTicketsPerTech = if ($config.PSObject.Properties.Name -contains 'TopTicketsPerTech') { [int]$config.TopTicketsPerTech } else { 10 }
$topTimeTicketsPerTech = if ($config.PSObject.Properties.Name -contains 'TopTimeTicketsPerTech') { [int]$config.TopTimeTicketsPerTech } else { 5 }

# Status display order
$statusOrder = @("Customer Reply","Waiting on Customer","In Progress","Quote/Billing","Resolved")

# Open tickets config (current unresolved)
$openStatuses = @()
$longOpenDays = 14
if ($config.PSObject.Properties.Name -contains 'OpenTickets' -and $config.OpenTickets) {
    if ($config.OpenTickets.PSObject.Properties.Name -contains 'Statuses') { $openStatuses = @($config.OpenTickets.Statuses) }
    if ($config.OpenTickets.PSObject.Properties.Name -contains 'LongOpenDays') { $longOpenDays = [int]$config.OpenTickets.LongOpenDays }
}
if (-not $openStatuses -or $openStatuses.Count -eq 0) {
    $openStatuses = @($statusOrder + @("Open")) | Select-Object -Unique
}

# Stale thresholds
$staleNoUpdateHours = 48
$staleCustomerReplyHours = 8
$staleWaitingOnCustomerDays = 5
if ($config.PSObject.Properties.Name -contains 'Stale' -and $config.Stale) {
    if ($config.Stale.PSObject.Properties.Name -contains 'NoUpdateHours') { $staleNoUpdateHours = [int]$config.Stale.NoUpdateHours }
    if ($config.Stale.PSObject.Properties.Name -contains 'CustomerReplyHours') { $staleCustomerReplyHours = [int]$config.Stale.CustomerReplyHours }
    if ($config.Stale.PSObject.Properties.Name -contains 'WaitingOnCustomerDays') { $staleWaitingOnCustomerDays = [int]$config.Stale.WaitingOnCustomerDays }
}

# Time logging config
$timeLoggingEnabled = $true
$timeSource = "ticket_timers"
$timePerPage = 200
$timeMaxPages = 10
if ($config.PSObject.Properties.Name -contains 'TimeLogging' -and $config.TimeLogging) {
    if ($config.TimeLogging.PSObject.Properties.Name -contains 'Enable')  { $timeLoggingEnabled = [bool]$config.TimeLogging.Enable }
    if ($config.TimeLogging.PSObject.Properties.Name -contains 'Source')  { $timeSource = [string]$config.TimeLogging.Source }
    if ($config.TimeLogging.PSObject.Properties.Name -contains 'PerPage') { $timePerPage = [int]$config.TimeLogging.PerPage }
    if ($config.TimeLogging.PSObject.Properties.Name -contains 'MaxPages'){ $timeMaxPages = [int]$config.TimeLogging.MaxPages }
}

# Hard tickets config (High priority + long open age)
$hardPriorities = @("High","Urgent","Emergency")
$hardMinAgeDays = 3
$hardPerTech = 0
$hardTopLevel = 0
if ($config.PSObject.Properties.Name -contains 'HardTickets' -and $config.HardTickets) {
    if ($config.HardTickets.PSObject.Properties.Name -contains 'Priorities') { $hardPriorities = @($config.HardTickets.Priorities) }
    if ($config.HardTickets.PSObject.Properties.Name -contains 'MinAgeDays') { $hardMinAgeDays = [int]$config.HardTickets.MinAgeDays }
    if ($config.HardTickets.PSObject.Properties.Name -contains 'PerTech') { $hardPerTech = [int]$config.HardTickets.PerTech }
    if ($config.HardTickets.PSObject.Properties.Name -contains 'TopLevel') { $hardTopLevel = [int]$config.HardTickets.TopLevel }
}
if ([string]::IsNullOrWhiteSpace($subdomain)) { throw "Config.Subdomain is empty." }
if ([string]::IsNullOrWhiteSpace($apiKey) -or $apiKey -like "PUT_*" -or $apiKey -like "PASTE_*") {
    throw "Syncro API key is not set. Set env var SYNCRO_API_KEY (recommended) or Config.ApiKey."
}

$baseDir = Split-Path -Parent $reportPath
if (!(Test-Path $baseDir)) { New-Item -ItemType Directory -Path $baseDir -Force | Out-Null }

$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $baseDir "Run_$runStamp.log"
$sampleTicketPath = Join-Path $baseDir "SampleTicket.json"
$sampleTimerPath  = Join-Path $baseDir "SampleTicketTimer.json"
$commentBodiesPath = Join-Path $baseDir "TechCommentBodies.json"
$techSummaryPath = Join-Path $baseDir "TechSummaries.txt"
$htmlReportPath = Join-Path $baseDir "LatestReport.html"
$htmlSummaryDir = $baseDir
$logoTarget = Join-Path $baseDir "logo.png"
foreach ($logoSource in @(
    (Join-Path $projectRoot "assets\BIGFOOT_WHITE_B200.png"),
    (Join-Path $scriptDir "BIGFOOT_WHITE_B200.png"),
    (Join-Path (Split-Path -Parent $baseDir) "BIGFOOT_WHITE_B200.png")
)) {
    if (Test-Path $logoSource) {
        Copy-Item -Path $logoSource -Destination $logoTarget -Force
        break
    }
}

Start-Transcript -Path $logPath | Out-Null

# -----------------------------
# Helpers
# -----------------------------
function Has-Prop($obj, $name) { ($null -ne $obj) -and ($null -ne $obj.PSObject.Properties[$name]) }
function Get-Prop($obj, $name) { if (Has-Prop $obj $name) { $obj.PSObject.Properties[$name].Value } else { $null } }
function Parse-Utc($value) { if ($null -eq $value) { return $null }; try { ([datetime]$value).ToUniversalTime() } catch { $null } }
function Minutes-ToHHMM([int]$mins) { $ts=[TimeSpan]::FromMinutes($mins); ('{0:00}:{1:00}' -f [int]$ts.TotalHours,$ts.Minutes) }
function Html-Encode([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($s)
}
function Slugify([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "tab" }
    $t = $s.ToLowerInvariant()
    $t = [regex]::Replace($t, "[^a-z0-9]+", "-")
    $t = $t.Trim("-")
    if (-not $t) { return "tab" }
    return $t
}

function Normalize-Text([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    return $s.Trim().ToLowerInvariant()
}

function Get-TechColor([string]$name) {
    $palette = @("#1d4ed8","#16a34a","#dc2626","#ea580c","#7c3aed","#0f766e","#be123c","#0ea5e9","#84cc16","#f59e0b")
    if ([string]::IsNullOrWhiteSpace($name)) { return "#94a3b8" }
    $hash = 0
    foreach ($ch in $name.ToCharArray()) {
        $hash = ($hash * 31 + [int][char]$ch) % 2147483647
    }
    $idx = $hash % $palette.Count
    return $palette[$idx]
}

function Get-StatusColor([string]$status) {
    switch ([string]$status) {
        "Resolved" { return "#22c55e" }
        "Customer Reply" { return "#ef4444" }
        "Waiting on Customer" { return "#f97316" }
        "Waiting on Supplier" { return "#f59e0b" }
        "Waiting for Parts" { return "#f59e0b" }
        "In Progress" { return "#3b82f6" }
        "Scheduled" { return "#0ea5e9" }
        "Quote/Billing" { return "#8b5cf6" }
        "Escalation" { return "#be123c" }
        "Open" { return "#64748b" }
        default { return "#64748b" }
    }
}

function Get-TicketUrl([string]$sub, $ticketId) {
    if ([string]::IsNullOrWhiteSpace($sub) -or -not $ticketId) { return $null }
    return "https://$sub.syncromsp.com/tickets/$ticketId"
}

function Clean-CommentText([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $lines = $text -split "`r`n|`n|`r"
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($lineRaw in $lines) {
        $line = $lineRaw.Trim()
        if (-not $line) { continue }
        $low = $line.ToLowerInvariant()
        if ($low -like "sent from *") { continue }
        if ($low -like "from:*") { continue }
        if ($low -like "sent:*") { continue }
        if ($low -like "to:*") { continue }
        if ($low -like "subject:*") { continue }
        if ($low -like "cc:*") { continue }
        if ($low -like "bcc:*") { continue }
        if ($low -like "importance:*") { continue }
        if ($low -like "on * wrote:*") { continue }
        if ($low -like "---*") { continue }
        if ($low -like "__*") { continue }
        if ($low -like "kind regards*") { continue }
        if ($low -like "regards*") { continue }
        if ($low -like "hi*") {
            $line = $line -replace '^(hi|hello|hey|good morning|good afternoon|good evening)[,:\s]+', ''
            $line = $line.Trim()
            if (-not $line) { continue }
            $low = $line.ToLowerInvariant()
        }
        if ($low -like "please consider the environment*") { continue }
        if ($low -like "this e-mail is confidential*") { continue }
        if ($low -like "this email is confidential*") { continue }
        if ($low -like "this e-mail and any attachments are confidential*") { continue }
        if ($low -like "this email, including any attachments, is strictly confidential*") { continue }
        if ($low -like "the information transmitted*") { continue }
        if ($low -like "any views or opinions presented*") { continue }
        if ($low -like "if you are not the intended recipient*") { continue }
        if ($low -like "please do not reply*") { continue }
        if ($low -like "please notify us immediately*") { continue }
        if ($low -like "please review our privacy policy*") { continue }
        if ($low -like "registered office*") { continue }
        if ($low -like "company limited by guarantee*") { continue }
        if ($low -like "company registered*") { continue }
        if ($low -like "registered in england*") { continue }
        if ($low -like "company no.*") { continue }
        if ($low -like "vat registration*") { continue }
        if ($low -like "internet communications are not secure*") { continue }
        if ($low -like "email communications are not secure*") { continue }
        if ($low -like "this footnote*") { continue }
        if ($low -like "please log all maintenance issues*") { continue }
        if ($low -like "disclaimer*") { continue }
        if ($low -like "[embedded image]*") { continue }
        if ($low -like "view your account profile online*") { continue }
        if ($low -like "want to tell us about our service*") { continue }
        if ($low -like "sent with care*") { continue }
        if ($low -like "unsubscribe*") { continue }

        if ($low -match '^(tel|t|phone|mob|mobile|fax|f|e|email|w|web|www)[:\s]') { continue }

        $line = $line -replace '\[embedded image\]', ''
        $line = $line -replace '\s+', ' '
        # strip URLs
        $line = [regex]::Replace($line, "https?://\S+", "").Trim()
        if ($line -match '^[\W_]+$') { continue }
        if ($line.Length -lt 6) { continue }
        if ($line) { $out.Add($line) }
    }
    return ($out -join " ")
}

function Get-ThemeCounts([string[]]$texts, [hashtable]$themes) {
    $counts = @{}
    foreach ($k in $themes.Keys) { $counts[$k] = 0 }
    foreach ($t in $texts) {
        $low = $t.ToLowerInvariant()
        foreach ($k in $themes.Keys) {
            foreach ($kw in $themes[$k]) {
                if ($low -like "*$kw*") { $counts[$k]++ ; break }
            }
        }
    }
    return $counts
}

function Get-TechNameFromTicket($t) {
    $u = Get-Prop $t 'user'
    if ($u -and (Has-Prop $u 'full_name') -and $u.full_name) { return [string]$u.full_name }
    return "Unassigned"
}

function Get-SyncroPaged {
    param(
        [Parameter(Mandatory=$true)][string]$UrlBase,
        [Parameter(Mandatory=$true)][string]$ItemsKey,
        [int]$MaxPages = 9999
    )

    $all = New-Object System.Collections.Generic.List[object]
    $page = 1
    $totalPages = 1

    do {
        $url = if ($UrlBase -match '\?') { "$UrlBase&page=$page" } else { "$UrlBase?page=$page" }
        Write-Host "GET $url"
        $resp = Invoke-RestMethod -Method GET -Uri $url

        $items = Get-Prop $resp $ItemsKey
        if ($items) { foreach ($i in $items) { $all.Add($i) } }

        $meta = Get-Prop $resp 'meta'
        $totalPages = 1
        if ($meta -and (Has-Prop $meta 'total_pages') -and $meta.total_pages) { $totalPages = [int]$meta.total_pages }

        $page++
        if ($page -gt $MaxPages) { break }
    } while ($page -le $totalPages)

    return $all
}

# -----------------------------
# Window: report period
# -----------------------------
$windowLocalStart = $null
$windowLocalEnd = $null
if ($windowMode -eq "LastWorkWeek") {
    $nowLocal = Get-Date
    $todayLocal = $nowLocal.Date
    $daysSinceMonday = (([int]$todayLocal.DayOfWeek + 6) % 7)
    $currentMonday = $todayLocal.AddDays(-$daysSinceMonday)
    $windowLocalStart = $currentMonday.AddDays(-7)
    $windowLocalEnd = $currentMonday.AddDays(-2)
    $startUtc = $windowLocalStart.ToUniversalTime()
    $endUtc = $windowLocalEnd.ToUniversalTime()
} else {
    $endUtc = if ($windowEndNow -and -not $windowEndAtUtcMidnight) { (Get-Date).ToUniversalTime() } else { (Get-Date).ToUniversalTime().Date }
    $startUtc = $endUtc.AddDays(-$daysBack)
}
if ($windowLocalStart -and $windowLocalEnd) {
    Write-Host "StartLocal: $($windowLocalStart.ToString('o'))"
    Write-Host "EndLocal:   $($windowLocalEnd.ToString('o'))"
}
Write-Host "StartUtc:   $($startUtc.ToString('o'))"
Write-Host "EndUtc:     $($endUtc.ToString('o'))"

$reportTitle = if ($windowMode -eq "LastWorkWeek") { "Weekly Technician Summary (Syncro)" } else { "Technician Summary (Syncro)" }

# -----------------------------
# Tickets updated since startUtc
# -----------------------------
$sinceIso = $startUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
$ticketsUrlBase = "https://$subdomain.syncromsp.com/api/v1/tickets?api_key=$apiKey&since_updated_at=$sinceIso"
$tickets = Get-SyncroPaged -UrlBase $ticketsUrlBase -ItemsKey "tickets"
Write-Host "Pulled tickets (raw): $($tickets.Count)"

if ($tickets.Count -gt 0) {
    ($tickets | Select-Object -First 1 | ConvertTo-Json -Depth 12) | Out-File $sampleTicketPath -Encoding UTF8
    Write-Host "Sample ticket saved to: $sampleTicketPath"
}

$ticketsInWindow = $tickets | Where-Object {
    $u = Parse-Utc (Get-Prop $_ 'updated_at')
    $u -and $u -ge $startUtc -and $u -lt $endUtc
}
Write-Host "Tickets updated in window: $(@($ticketsInWindow).Count)"

$ticketRows = $ticketsInWindow | ForEach-Object {
    [pscustomobject]@{
        Tech      = Get-TechNameFromTicket $_
        TicketId  = (Get-Prop $_ 'id')
        Number    = (Get-Prop $_ 'number')
        Subject   = (Get-Prop $_ 'subject')
        Status    = (Get-Prop $_ 'status')
        Priority  = (Get-Prop $_ 'priority')
        Customer  = (Get-Prop $_ 'customer_business_then_name')
        CreatedAt = (Get-Prop $_ 'created_at')
        UpdatedAt = (Get-Prop $_ 'updated_at')
    }
}

$ticketsCreated = $ticketRows | Where-Object {
    $c = Parse-Utc $_.CreatedAt
    $c -and $c -ge $startUtc -and $c -lt $endUtc
}

# -----------------------------
# Current open tickets (all unresolved)
# -----------------------------
$nowUtc = (Get-Date).ToUniversalTime()
$openTicketsRaw = New-Object System.Collections.Generic.List[object]
foreach ($st in $openStatuses) {
    if ([string]::IsNullOrWhiteSpace($st)) { continue }
    $stEnc = [uri]::EscapeDataString($st)
    $openUrlBase = "https://$subdomain.syncromsp.com/api/v1/tickets?api_key=$apiKey&status=$stEnc"
    $rows = Get-SyncroPaged -UrlBase $openUrlBase -ItemsKey "tickets"
    if ($rows) { foreach ($r in $rows) { $openTicketsRaw.Add($r) } }
}
Write-Host "Pulled open tickets (raw, by status): $($openTicketsRaw.Count)"

$openById = @{}
foreach ($t in $openTicketsRaw) {
    $id = Get-Prop $t 'id'
    if ($id -and -not $openById.ContainsKey($id)) { $openById[$id] = $t }
}
$openTickets = $openById.Values

$openTicketRows = $openTickets | ForEach-Object {
    $status = [string](Get-Prop $_ 'status')
    if ($status -eq "Resolved") { return $null }
    $created = Parse-Utc (Get-Prop $_ 'created_at')
    if (-not $created) { $created = Parse-Utc (Get-Prop $_ 'updated_at') }
    $ageDays = if ($created) { [math]::Round((($nowUtc - $created).TotalDays),1) } else { $null }
    [pscustomobject]@{
        Tech      = Get-TechNameFromTicket $_
        TicketId  = (Get-Prop $_ 'id')
        Number    = (Get-Prop $_ 'number')
        Subject   = (Get-Prop $_ 'subject')
        Status    = $status
        Priority  = (Get-Prop $_ 'priority')
        Customer  = (Get-Prop $_ 'customer_business_then_name')
        CreatedAt = (Get-Prop $_ 'created_at')
        UpdatedAt = (Get-Prop $_ 'updated_at')
        AgeDays   = $ageDays
        LongOpen  = ($null -ne $ageDays -and $ageDays -ge $longOpenDays)
    }
} | Where-Object { $_ }

$openByTech = $openTicketRows | Group-Object Tech | Sort-Object Name
$openByTechMap = @{}
foreach ($g in $openByTech) { $openByTechMap[$g.Name] = $g.Group }

$openLastWeekUnresolved = $openTicketRows | Where-Object {
    $c = Parse-Utc $_.CreatedAt
    $c -and $c -ge $startUtc -and $c -lt $endUtc
}

# -----------------------------
# Comment bodies for summaries (visible only)
# -----------------------------
$commentBodiesByTech = @{}
foreach ($t in $ticketsInWindow) {
    $tech = Get-TechNameFromTicket $t
    $comments = Get-Prop $t 'comments'
    if (-not $comments) { continue }

    foreach ($c in $comments) {
        $hidden = $false
        if (Has-Prop $c 'hidden') { $hidden = [bool]$c.hidden }
        if ($hidden) { continue }

        $body = Get-Prop $c 'body'
        if ([string]::IsNullOrWhiteSpace($body)) { continue }

        if (-not $commentBodiesByTech.ContainsKey($tech)) {
            $commentBodiesByTech[$tech] = New-Object System.Collections.Generic.List[string]
        }

        $commentBodiesByTech[$tech].Add([string]$body)
    }
}

# Write raw comment bodies for offline summarization
$commentBodiesByTech | ConvertTo-Json -Depth 5 | Out-File -FilePath $commentBodiesPath -Encoding UTF8
Write-Host "Comment bodies written to: $commentBodiesPath"

# -----------------------------
# Tech summaries (automated)
# -----------------------------
$themes = @{
    "Email/Outlook" = @("email","outlook","mailbox","shared","bounce","undelivered","smtp","bcc","inbox")
    "Accounts/Access" = @("account","access","login","password","permission","license","azure","user","leaver","starter")
    "Syncro/Agents" = @("syncro","agent","installer","accept","device","asset")
    "Hardware/Devices" = @("laptop","printer","camera","microphone","keyboard","monitor","pc","device")
    "Servers/Network" = @("server","vpn","dns","rds","network","slow","latency","disk","backup")
    "Security" = @("phish","malicious","vulnerability","blocked","security","alert","attack")
    "Quoting/Billing" = @("quote","estimate","billing","invoice","charge","order")
    "Scheduling/Follow-up" = @("follow up","checking","let me know","schedule","out of hours","tonight")
}

$summaryLines = New-Object System.Collections.Generic.List[string]
$summaryLines.Add("Technician Summaries (Automated)")
if ($windowLocalStart -and $windowLocalEnd) {
    $summaryLines.Add("Window (Local): $($windowLocalStart.ToString('yyyy-MM-dd HH:mm')) -> $($windowLocalEnd.ToString('yyyy-MM-dd HH:mm'))")
}
$summaryLines.Add("Window (UTC): $($startUtc.ToString('yyyy-MM-dd HH:mm')) -> $($endUtc.ToString('yyyy-MM-dd HH:mm'))")
$summaryLines.Add("Generated (local): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$summaryLines.Add("")

foreach ($tech in $commentBodiesByTech.Keys) {
    $rawList = $commentBodiesByTech[$tech]
    $cleaned = @()
    $seen = @{}
    foreach ($t in $rawList) {
        $c = Clean-CommentText $t
        if ($c) {
            if (-not $seen.ContainsKey($c)) {
                $seen[$c] = $true
                $cleaned += $c
            }
        }
    }

    $summaryLines.Add("=== $tech ===")
    $summaryLines.Add(("Messages analyzed: {0}" -f @($cleaned).Count))

    if (@($cleaned).Count -eq 0) {
        $summaryLines.Add("No visible comment bodies found.")
        $summaryLines.Add("")
        continue
    }

    $counts = Get-ThemeCounts -texts $cleaned -themes $themes
    $topThemes =
        $counts.GetEnumerator() |
        Sort-Object Value -Descending |
        Where-Object { $_.Value -gt 0 } |
        Select-Object -First 4

    if ($topThemes) {
        $topTxt = ($topThemes | ForEach-Object { "$($_.Key) ($($_.Value))" }) -join ", "
        $summaryLines.Add("Top themes: $topTxt")
    }

    # Build a short narrative using top themes
    $themeNames = @($topThemes | ForEach-Object { $_.Key })
    if ($themeNames.Count -gt 0) {
        $summaryLines.Add(("Focus areas included {0}." -f ($themeNames -join ", ")))
    }

    # Provide up to 3 representative cleaned snippets
    $summaryLines.Add("Representative updates:")
    $snips =
        $cleaned |
        Where-Object { $_.Length -ge 40 -and $_.Length -le 240 } |
        Sort-Object Length -Descending |
        Select-Object -First 3
    if (-not $snips -or @($snips).Count -eq 0) {
        $snips = $cleaned | Select-Object -First 3
    }
    foreach ($s in $snips) {
        $summaryLines.Add(" - $s")
    }
    $summaryLines.Add("")
}

($summaryLines -join "`r`n") | Out-File -FilePath $techSummaryPath -Encoding UTF8
Write-Host "Tech summaries written to: $techSummaryPath"

# -----------------------------
# Stale
# -----------------------------
$noUpdateCutoff = $nowUtc.AddHours(-$staleNoUpdateHours)
$customerReplyCutoff = $nowUtc.AddHours(-$staleCustomerReplyHours)
$waitingCustomerCutoff = $nowUtc.AddDays(-$staleWaitingOnCustomerDays)

$staleCustomerReply = $ticketRows | Where-Object { $_.Status -eq "Customer Reply" -and (Parse-Utc $_.UpdatedAt) -lt $customerReplyCutoff }
$staleWaitingOnCustomer = $ticketRows | Where-Object { $_.Status -eq "Waiting on Customer" -and (Parse-Utc $_.UpdatedAt) -lt $waitingCustomerCutoff }
$staleNoUpdate = $ticketRows | Where-Object { (Parse-Utc $_.UpdatedAt) -lt $noUpdateCutoff -and $_.Status -ne "Resolved" }

# -----------------------------
# Hard tickets (High priority + long open age)
# -----------------------------
$hardPrioritySet = @{}
foreach ($p in $hardPriorities) {
    $k = Normalize-Text $p
    if ($k) { $hardPrioritySet[$k] = $true }
}

$hardCandidates = $ticketRows | ForEach-Object {
    $prioKey = Normalize-Text $_.Priority
    if (-not $hardPrioritySet.ContainsKey($prioKey)) { return $null }
    if ($_.Status -eq "Resolved") { return $null }

    $created = Parse-Utc $_.CreatedAt
    if (-not $created) { $created = Parse-Utc $_.UpdatedAt }
    if (-not $created) { return $null }

    $ageDays = [math]::Round((($nowUtc - $created).TotalDays),1)
    if ($ageDays -lt $hardMinAgeDays) { return $null }

    [pscustomobject]@{
        Tech = $_.Tech
        TicketId = $_.TicketId
        Number = $_.Number
        Subject = $_.Subject
        Status = $_.Status
        Priority = $_.Priority
        Customer = $_.Customer
        CreatedAt = $_.CreatedAt
        AgeDays = $ageDays
    }
} | Where-Object { $_ }

$hardTopLevelTickets = $hardCandidates | Sort-Object AgeDays -Descending | Select-Object

$hardByTech = @{}
$hardGroups = $hardCandidates | Group-Object Tech
foreach ($g in $hardGroups) {
    $hardByTech[$g.Name] = $g.Group | Sort-Object AgeDays -Descending | Select-Object
}

# -----------------------------
# Time logged from /ticket_timers
# -----------------------------
$timeTotalsByTech = @{}
$avgMinsPerTicketByTech = @{}
$topTimeByTech = @{}
$timerWindowRows = @()

function Get-TechFromTimer($tm) {
    # common: timer.user.full_name (guess based on Syncro patterns)
    $u = Get-Prop $tm 'user'
    if ($u -and (Has-Prop $u 'full_name') -and $u.full_name) { return [string]$u.full_name }

    # sometimes: 'tech' string
    $tech = Get-Prop $tm 'tech'
    if ($tech) { return [string]$tech }

    # sometimes: user_name
    $un = Get-Prop $tm 'user_name'
    if ($un) { return [string]$un }

    $uid = Get-Prop $tm 'user_id'
    if ($uid) { return "UserId:$uid" }

    return "Unknown"
}

function Get-TimerWhenUtc($tm) {
    # Try likely fields in order
    foreach ($k in @('updated_at','stopped_at','ended_at','created_at','started_at')) {
        $d = Parse-Utc (Get-Prop $tm $k)
        if ($d) { return $d }
    }
    return $null
}

function Get-TimerMinutes($tm) {
    # Try common fields
    foreach ($k in @('minutes','duration_minutes','elapsed_minutes','time_spent_minutes','duration')) {
        $v = Get-Prop $tm $k
        if ($null -ne $v) { try { return [int][math]::Round([double]$v,0) } catch {} }
    }

    # If start/stop timestamps exist, compute minutes
    $start = Parse-Utc (Get-Prop $tm 'started_at')
    if (-not $start) { $start = Parse-Utc (Get-Prop $tm 'start_at') }
    $stop  = Parse-Utc (Get-Prop $tm 'stopped_at')
    if (-not $stop) { $stop = Parse-Utc (Get-Prop $tm 'ended_at') }

    if ($start -and $stop -and $stop -ge $start) {
        return [int][math]::Round(($stop - $start).TotalMinutes,0)
    }

    return 0
}

function Get-TimerBillable($tm) {
    foreach ($k in @('billable','is_billable','billable?')) {
        $v = Get-Prop $tm $k
        if ($null -ne $v) { return [bool]$v }
    }
    return $false
}

function Get-TimerTicketNumber($tm) {
    $t = Get-Prop $tm 'ticket'
    if ($t -and (Has-Prop $t 'number')) { return $t.number }
    $v = Get-Prop $tm 'ticket_number'
    if ($v) { return $v }
    # sometimes only ticket_id; we’ll use it as a key if needed
    return $null
}

function Get-TimerTicketId($tm) {
    foreach ($k in @('ticket_id')) {
        $v = Get-Prop $tm $k
        if ($v) { return $v }
    }
    $t = Get-Prop $tm 'ticket'
    if ($t -and (Has-Prop $t 'id')) { return $t.id }
    return $null
}

function Get-TimerTicketSubject($tm) {
    $t = Get-Prop $tm 'ticket'
    if ($t -and (Has-Prop $t 'subject')) { return [string]$t.subject }
    return $null
}

if ($timeLoggingEnabled -and $timeSource -eq "ticket_timers") {
    $timersUrlBase = "https://$subdomain.syncromsp.com/api/v1/ticket_timers?api_key=$apiKey&per_page=$timePerPage"
    $timers = Get-SyncroPaged -UrlBase $timersUrlBase -ItemsKey "ticket_timers" -MaxPages $timeMaxPages
    Write-Host "Pulled ticket_timers (raw, capped): $($timers.Count)"

    if ($timers.Count -gt 0) {
        ($timers | Select-Object -First 1 | ConvertTo-Json -Depth 12) | Out-File $sampleTimerPath -Encoding UTF8
        Write-Host "Sample ticket timer saved to: $sampleTimerPath"
    }

    $timerWindowRows = $timers | ForEach-Object {
        $when = Get-TimerWhenUtc $_
        $mins = Get-TimerMinutes $_
        [pscustomobject]@{
            Tech = Get-TechFromTimer $_
            WhenUtc = $when
            Minutes = $mins
            Billable = Get-TimerBillable $_
            TicketId = Get-TimerTicketId $_
            TicketNumber = Get-TimerTicketNumber $_
            TicketSubject = Get-TimerTicketSubject $_
        }
    } | Where-Object { $_.WhenUtc -and $_.WhenUtc -ge $startUtc -and $_.WhenUtc -lt $endUtc -and $_.Minutes -gt 0 }

    Write-Host "Ticket timers in window (with minutes): $(@($timerWindowRows).Count)"

    $byTech = $timerWindowRows | Group-Object Tech
    foreach ($g in $byTech) {
        $tech = $g.Name
        $totalMins = [int](($g.Group | Measure-Object -Property Minutes -Sum).Sum)
        $billMins  = [int](($g.Group | Where-Object { $_.Billable -eq $true } | Measure-Object -Property Minutes -Sum).Sum)
        $nonBillMins = $totalMins - $billMins

        # Distinct tickets touched: prefer number else id
        $ticketKeys = $g.Group | ForEach-Object {
            if ($_.TicketNumber) { "num:$($_.TicketNumber)" } elseif ($_.TicketId) { "id:$($_.TicketId)" } else { $null }
        } | Where-Object { $_ } | Select-Object -Unique
        $distinctTickets = @($ticketKeys).Count
        $avg = if ($distinctTickets -gt 0) { [math]::Round($totalMins / $distinctTickets, 1) } else { 0 }

        $topTickets =
            $g.Group |
            Group-Object { if ($_.TicketNumber) { $_.TicketNumber } elseif ($_.TicketId) { $_.TicketId } else { "UnknownTicket" } } |
            ForEach-Object {
                $sumM = [int](($_.Group | Measure-Object -Property Minutes -Sum).Sum)
                $first = $_.Group | Select-Object -First 1
                [pscustomobject]@{
                    Key = $_.Name
                    Minutes = $sumM
                    TicketNumber = $first.TicketNumber
                    TicketSubject = $first.TicketSubject
                }
            } |
            Sort-Object Minutes -Descending |
            Select-Object -First $topTimeTicketsPerTech

        $timeTotalsByTech[$tech] = [pscustomobject]@{
            TotalMins = $totalMins
            BillableMins = $billMins
            NonBillableMins = $nonBillMins
        }
        $avgMinsPerTicketByTech[$tech] = $avg
        $topTimeByTech[$tech] = $topTickets
    }
}

$totalClosed = @($ticketRows | Where-Object { $_.Status -eq "Resolved" }).Count
$totalOpen = @($openTicketRows).Count
$totalLongOpen = @($openTicketRows | Where-Object { $_.LongOpen }).Count
$totalTimeMins = 0
$totalBillableMins = 0
$totalNonBillableMins = 0
foreach ($t in $timeTotalsByTech.Values) {
    $totalTimeMins += [int]$t.TotalMins
    $totalBillableMins += [int]$t.BillableMins
    $totalNonBillableMins += [int]$t.NonBillableMins
}

# -----------------------------
# Report
# -----------------------------
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add($reportTitle)
if ($windowLocalStart -and $windowLocalEnd) {
    $lines.Add("Window (Local): $($windowLocalStart.ToString('yyyy-MM-dd HH:mm')) -> $($windowLocalEnd.ToString('yyyy-MM-dd HH:mm'))")
}
$lines.Add("Window (UTC): $($startUtc.ToString('yyyy-MM-dd HH:mm')) -> $($endUtc.ToString('yyyy-MM-dd HH:mm'))")
$lines.Add("Total tickets updated in window: $(@($ticketRows).Count)")
$lines.Add("Generated (local): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$lines.Add("=== Executive Summary ===")
$lines.Add(("Tickets updated: {0} | Closed (Resolved): {1}" -f @($ticketRows).Count, $totalClosed))
$lines.Add(("Current open (unresolved): {0} | Long open >= {1} days: {2}" -f $totalOpen, $longOpenDays, $totalLongOpen))
if ($timeLoggingEnabled) {
    $lines.Add(("Time logged in window: {0} (Billable {1}, Non-billable {2})" -f (Minutes-ToHHMM $totalTimeMins), (Minutes-ToHHMM $totalBillableMins), (Minutes-ToHHMM $totalNonBillableMins)))
}
$lines.Add(("Hard tickets (High priority, open >= {0} days): {1}" -f $hardMinAgeDays, @($hardCandidates).Count))
if ($hardTopLevelTickets -and @($hardTopLevelTickets).Count -gt 0) {
    $lines.Add("All hard tickets:")
    foreach ($ht in $hardTopLevelTickets) {
        $label = if ($ht.Number) { "#" + $ht.Number } elseif ($ht.TicketId) { "TicketId:" + $ht.TicketId } else { "UnknownTicket" }
        $prio = if ($ht.Priority) { $ht.Priority } else { "Priority?" }
        $subj = if ($ht.Subject) { $ht.Subject } else { "" }
        $cust = if ($ht.Customer) { $ht.Customer } else { "Unknown Customer" }
        $lines.Add((" - {0} ({1}, {2}d open) {3}— {4}" -f $label, $prio, $ht.AgeDays, $subj, $cust))
    }
}
$lines.Add("")

$lines.Add("=== Stale / Needs Attention ===")
$lines.Add(("No update in last {0}h (and not Resolved): {1}" -f $staleNoUpdateHours, @($staleNoUpdate).Count))
$lines.Add(("Customer Reply older than {0}h: {1}" -f $staleCustomerReplyHours, @($staleCustomerReply).Count))
$lines.Add(("Waiting on Customer older than {0} days: {1}" -f $staleWaitingOnCustomerDays, @($staleWaitingOnCustomer).Count))
$lines.Add("")

$lines.Add("=== Per Technician ===")
$lines.Add("")

$byTechTickets = $ticketRows | Group-Object Tech | Sort-Object Name
foreach ($grp in $byTechTickets) {
    $tech = $grp.Name
    $closed = @($grp.Group | Where-Object { $_.Status -eq "Resolved" }).Count

    $lines.Add("=== $tech ===")
    $lines.Add("Tickets updated: $(@($grp.Group).Count)")
    $lines.Add("Tickets closed (Resolved): $closed")
    $statusCounts = $grp.Group | Group-Object Status | Sort-Object Name
    if ($statusCounts) {
        $statusLine = ($statusCounts | ForEach-Object { "$($_.Name): $($_.Count)" }) -join " | "
        $lines.Add("Status breakdown: $statusLine")
    }

    if ($timeTotalsByTech.ContainsKey($tech)) {
        $tTot = $timeTotalsByTech[$tech]
        $lines.Add(("Time logged in window: {0} (Billable {1}, Non-billable {2})" -f `
            (Minutes-ToHHMM $tTot.TotalMins), (Minutes-ToHHMM $tTot.BillableMins), (Minutes-ToHHMM $tTot.NonBillableMins)))
        $lines.Add(("Avg mins per ticket touched (from ticket timers): {0}" -f $avgMinsPerTicketByTech[$tech]))

        $topTime = $topTimeByTech[$tech]
        if ($topTime -and @($topTime).Count -gt 0) {
            $lines.Add("Top tickets by time (window):")
            foreach ($tt in $topTime) {
                $label = if ($tt.TicketNumber) { "#$($tt.TicketNumber)" } else { "TicketId:$($tt.Key)" }
                $subj  = if ($tt.TicketSubject) { $tt.TicketSubject } else { "" }
                $lines.Add((" - {0} {1} ({2})" -f $label, $subj, (Minutes-ToHHMM $tt.Minutes)))
            }
        }
    } elseif ($timeLoggingEnabled) {
        $lines.Add("Time logged in window: 0 (no timers found in window OR mapping differs)")
    }

    $allTickets = $grp.Group | Sort-Object { Parse-Utc $_.UpdatedAt } -Descending
    foreach ($t in $allTickets) {
        $cust = if ($t.Customer) { $t.Customer } else { "Unknown Customer" }
        $prio = if ($t.Priority) { $t.Priority } else { "" }
        $prioTxt = if ($prio) { "($prio) " } else { "" }
        $lines.Add((" - #{0} [{1}] {2}{3}— {4}" -f $t.Number, $t.Status, $prioTxt, $t.Subject, $cust))
    }

    $openTechTickets = if ($openByTechMap.ContainsKey($tech)) { $openByTechMap[$tech] } else { @() }
    $lines.Add(("Current open tickets (all unresolved): {0}" -f @($openTechTickets).Count))
    if (@($openTechTickets).Count -gt 0) {
        foreach ($ot in ($openTechTickets | Sort-Object AgeDays -Descending)) {
            $cust = if ($ot.Customer) { $ot.Customer } else { "Unknown Customer" }
            $prio = if ($ot.Priority) { $ot.Priority } else { "" }
            $prioTxt = if ($prio) { "($prio) " } else { "" }
            $ageTxt = if ($null -ne $ot.AgeDays) { "$($ot.AgeDays)d" } else { "Unknown age" }
            $flag = if ($ot.LongOpen) { " [LONG OPEN]" } else { "" }
            $lines.Add((" - #{0} [{1}] {2}{3}— {4} (Age {5}){6}" -f $ot.Number, $ot.Status, $prioTxt, $ot.Subject, $cust, $ageTxt, $flag))
        }
    }

    $lines.Add("")
}

# -----------------------------
# Append automated tech summaries
# -----------------------------
$lines.Add("=== Tech Summaries (Automated) ===")
if ($windowLocalStart -and $windowLocalEnd) {
    $lines.Add(("Window (Local): $($windowLocalStart.ToString('yyyy-MM-dd HH:mm')) -> $($windowLocalEnd.ToString('yyyy-MM-dd HH:mm'))"))
}
$lines.Add(("Window (UTC): $($startUtc.ToString('yyyy-MM-dd HH:mm')) -> $($endUtc.ToString('yyyy-MM-dd HH:mm'))"))
$lines.Add(("Generated (local): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"))
$lines.Add("")
$summaryArr = $summaryLines.ToArray()
$breakIdx = [array]::IndexOf($summaryArr, "")
$summaryBody =
    if ($breakIdx -ge 0 -and $breakIdx -lt ($summaryArr.Length - 1)) {
        $summaryArr[($breakIdx + 1)..($summaryArr.Length - 1)]
    } else {
        $summaryArr
    }
foreach ($l in $summaryBody) {
    $lines.Add($l)
}

($lines -join "`r`n") | Out-File -FilePath $reportPath -Encoding UTF8
Write-Host "Report written to: $reportPath"
Write-Host "Log written to: $logPath"

# Overall status counts (all techs)
$overallStatusCounts = $ticketRows | Group-Object Status
$overallStatusCounts =
    $overallStatusCounts |
    Sort-Object {
        $idx = $statusOrder.IndexOf([string]$_.Name)
        if ($idx -ge 0) { $idx } else { 999 }
    }, Name

function Write-SummaryPage {
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][string]$FileName,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Tickets,
        [Parameter(Mandatory=$true)][string]$Accent
    )

    $p = Join-Path $htmlSummaryDir $FileName
    $h = New-Object System.Collections.Generic.List[string]
    $h.Add("<!doctype html>")
    $h.Add("<html lang=`"en`">")
    $h.Add("<head>")
    $h.Add("  <meta charset=`"utf-8`" />")
    $h.Add("  <meta name=`"viewport`" content=`"width=device-width, initial-scale=1`" />")
    $h.Add("  <title>$(Html-Encode($Title))</title>")
    $h.Add("  <style>")
    $h.Add("    :root { --ink:#0f172a; --muted:#64748b; --panel:#ffffff; --outline:#e2e8f0; --accent:#0b4aa2; --accent-2:#0f6fb3; --bg1:#0b2a57; --bg2:#0b4aa2; --bg3:#0f6fb3; }")
    $h.Add("    body { margin:0; font-family:`"Manrope`",`"Avenir Next`",`"Trebuchet MS`",sans-serif; background:linear-gradient(140deg,var(--bg1),var(--bg2) 40%,var(--bg3) 85%); color:var(--ink); background-image:radial-gradient(rgba(255,255,255,0.06) 1px, transparent 1px); background-size:18px 18px; }")
    $h.Add("    .wrap { max-width:1200px; margin:0 auto; padding:28px 24px 40px; }")
    $h.Add("    .card { background:var(--panel); border-radius:16px; padding:16px 18px; box-shadow:0 10px 30px rgba(15,23,42,0.12); border:1px solid rgba(226,232,240,0.9); }")
    $h.Add("    .nav-btn { display:inline-flex; align-items:center; gap:8px; padding:8px 14px; border-radius:999px; background:rgba(255,255,255,0.95); border:1px solid rgba(226,232,240,0.9); color:#0f172a; text-decoration:none; font-weight:700; letter-spacing:0.02em; transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; }")
    $h.Add("    .nav-btn:hover { border-color:#94a3b8; box-shadow:0 8px 18px rgba(15,23,42,0.12); transform:translateY(-1px); }")
    $h.Add("    .accent-blue { border-color:#93c5fd; background:#eff6ff; }")
    $h.Add("    .accent-green { border-color:#86efac; background:#dcfce7; }")
    $h.Add("    .accent-red { border-color:#fca5a5; background:#fee2e2; }")
    $h.Add("    .accent-orange { border-color:#fdba74; background:#ffedd5; }")
    $h.Add("    .accent-yellow { border-color:#fde68a; background:#fef3c7; }")
    $h.Add("    .muted { color:#64748b; }")
    $h.Add("    .ticket { margin:10px 0; padding:10px 12px; border:1px solid rgba(226,232,240,0.9); border-radius:12px; background:#fff; box-shadow:0 6px 16px rgba(15,23,42,0.08); }")
    $h.Add("    .tech-head { margin:12px 0 6px 0; padding:8px 12px; border-radius:12px; color:#fff; font-weight:700; letter-spacing:0.02em; }")
    $h.Add("    .status-pill { display:inline-block; padding:2px 8px; border-radius:999px; font-size:11px; color:#fff; margin-left:6px; }")
    $h.Add("    a { color:var(--accent-2); text-decoration:none; }")
    $h.Add("  </style>")
    $h.Add("</head>")
    $h.Add("<body><div class=`"wrap`">")
    $h.Add("<div class=`"card $Accent`">")
    $h.Add("<div class=`"muted`"><a class=`"nav-btn`" href=`"LatestReport.html`">Return to Summary</a></div>")
    $h.Add("<h1>$(Html-Encode($Title))</h1>")
    if ($windowLocalStart -and $windowLocalEnd) {
        $h.Add("<div class=`"muted`">Window (Local): $(Html-Encode($windowLocalStart.ToString('yyyy-MM-dd HH:mm'))) → $(Html-Encode($windowLocalEnd.ToString('yyyy-MM-dd HH:mm')))</div>")
    }
    $h.Add("<div class=`"muted`">Window (UTC): $(Html-Encode($startUtc.ToString('yyyy-MM-dd HH:mm'))) → $(Html-Encode($endUtc.ToString('yyyy-MM-dd HH:mm')))</div>")
    $h.Add("<div class=`"muted`">Total: $(@($Tickets).Count)</div>")
    $h.Add("</div>")
    if (@($Tickets).Count -gt 0) {
        $byTech = $Tickets | Group-Object Tech | Sort-Object Name
        foreach ($g in $byTech) {
            $tech = $g.Name
            $color = Get-TechColor $tech
            $h.Add("<div class=`"tech-head`" style=`"background:$color`">$(Html-Encode($tech))</div>")
            foreach ($t in $g.Group) {
                $cust = if ($t.Customer) { $t.Customer } else { "Unknown Customer" }
                $prio = if ($t.Priority) { $t.Priority } else { "" }
                $prioTxt = if ($prio) { "($prio) " } else { "" }
                $statusColor = Get-StatusColor $t.Status
                $statusPill = "<span class=`"status-pill`" style=`"background:$statusColor`">$(Html-Encode($t.Status))</span>"
                $h.Add("<div class=`"ticket`" style=`"border-left:4px solid $color`">#$(Html-Encode($t.Number)) $statusPill $(Html-Encode($prioTxt + $t.Subject)) — $(Html-Encode($cust)) <span class=`"muted`">| Tech: $(Html-Encode($tech))</span></div>")
            }
        }
    } else {
        $h.Add("<div class=`"muted`">No tickets.</div>")
    }
    $h.Add("</div></body></html>")
    ($h -join "`r`n") | Out-File -FilePath $p -Encoding UTF8
}

# Per-tech open tickets page (current unresolved)
function Write-TechOpenPage {
    param(
        [Parameter(Mandatory=$true)][string]$Tech,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Tickets,
        [Parameter(Mandatory=$true)][string]$FileName
    )

    $p = Join-Path $htmlSummaryDir $FileName
    $h = New-Object System.Collections.Generic.List[string]
    $h.Add("<!doctype html>")
    $h.Add("<html lang=`"en`">")
    $h.Add("<head>")
    $h.Add("  <meta charset=`"utf-8`" />")
    $h.Add("  <meta name=`"viewport`" content=`"width=device-width, initial-scale=1`" />")
    $h.Add("  <title>$(Html-Encode($Tech)) - Open Tickets</title>")
    $h.Add("  <style>")
    $h.Add("    :root { --ink:#0f172a; --muted:#64748b; --panel:#ffffff; --outline:#e2e8f0; --accent:#0b4aa2; --accent-2:#0f6fb3; --bg1:#0b2a57; --bg2:#0b4aa2; --bg3:#0f6fb3; }")
    $h.Add("    body { margin:0; font-family:`"Manrope`",`"Avenir Next`",`"Trebuchet MS`",sans-serif; background:linear-gradient(140deg,var(--bg1),var(--bg2) 40%,var(--bg3) 85%); color:var(--ink); background-image:radial-gradient(rgba(255,255,255,0.06) 1px, transparent 1px); background-size:18px 18px; }")
    $h.Add("    .wrap { max-width:1200px; margin:0 auto; padding:28px 24px 40px; }")
    $h.Add("    .card { background:var(--panel); border-radius:16px; padding:16px 18px; box-shadow:0 10px 30px rgba(15,23,42,0.12); border:1px solid rgba(226,232,240,0.9); }")
    $h.Add("    .nav-btn { display:inline-flex; align-items:center; gap:8px; padding:8px 14px; border-radius:999px; background:rgba(255,255,255,0.95); border:1px solid rgba(226,232,240,0.9); color:#0f172a; text-decoration:none; font-weight:700; letter-spacing:0.02em; transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; }")
    $h.Add("    .nav-btn:hover { border-color:#94a3b8; box-shadow:0 8px 18px rgba(15,23,42,0.12); transform:translateY(-1px); }")
    $h.Add("    .muted { color:#e2e8f0; }")
    $h.Add("    .badge { display:inline-block; padding:2px 8px; border-radius:999px; background:#e2e8f0; font-size:12px; margin-left:6px; }")
    $h.Add("    .badge-long { background:#fecaca; border:1px solid #fca5a5; color:#7f1d1d; }")
    $h.Add("    .status-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(260px,1fr)); gap:16px; align-items:start; margin-top:12px; }")
    $h.Add("    .status-col { background:rgba(255,255,255,0.95); border:1px solid rgba(226,232,240,0.9); border-radius:14px; padding:12px; box-shadow:0 8px 18px rgba(15,23,42,0.08); }")
    $h.Add("    .status-col h4 { margin:0 0 10px 0; font-size:14px; letter-spacing:0.02em; }")
    $h.Add("    .status-list { max-height:520px; overflow:auto; padding-right:6px; }")
    $h.Add("    .status-open { border-color:#f59e0b; background:#fef3c7; }")
    $h.Add("    .status-waiting { border-color:#f97316; background:#ffedd5; }")
    $h.Add("    .status-customer { border-color:#ef4444; background:#fee2e2; }")
    $h.Add("    .status-resolved { border-color:#22c55e; background:#dcfce7; }")
    $h.Add("    .status-inprogress { border-color:#3b82f6; background:#dbeafe; }")
    $h.Add("    .status-quotebilling { border-color:#8b5cf6; background:#ede9fe; }")
    $h.Add("    .ticket-card { display:block; background:#ffffff; border:1px solid rgba(226,232,240,0.9); border-radius:12px; padding:10px 12px; margin:10px 0; color:inherit; text-decoration:none; box-shadow:0 6px 16px rgba(15,23,42,0.08); transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; }")
    $h.Add("    .ticket-card:hover { border-color:#94a3b8; box-shadow:0 10px 20px rgba(15,23,42,0.12); transform:translateY(-1px); }")
    $h.Add("  </style>")
    $h.Add("</head>")
    $h.Add("<body><div class=`"wrap`">")
    $techColor = Get-TechColor $Tech
    $h.Add("<div class=`"card`" style=`"background:#e6f2fb; color:#111827; border-color:#cfe2f3`">")
    $h.Add("<div><a class=`"nav-btn`" href=`"LatestReport.html`">Return to Summary</a></div>")
    $h.Add("<div style=`"background:$techColor; color:#ffffff; font-weight:700; padding:8px 10px; border-radius:8px; margin-top:8px;`">$(Html-Encode($Tech)) - Current Open Tickets</div>")
    $h.Add("<div style=`"background:#0077c0; color:#ffffff; padding:6px 10px; border-radius:8px; margin-top:6px; display:inline-block;`">Open: $(@($Tickets).Count) | $longOpenDays+ day open: $(@($Tickets | Where-Object { $_.LongOpen }).Count)</div>")
    $h.Add("</div>")

    if (@($Tickets).Count -gt 0) {
        $byStatus = $Tickets | Group-Object Status
        $byStatus =
            $byStatus |
            Sort-Object {
                $idx = $statusOrder.IndexOf([string]$_.Name)
                if ($idx -ge 0) { $idx } else { 999 }
            }, Name

        $h.Add("<div class=`"status-grid`">")
        foreach ($sg in $byStatus) {
            $statusName = [string]$sg.Name
            $statusClass = "status-open"
            switch ($statusName) {
                "Resolved" { $statusClass = "status-resolved" }
                "Customer Reply" { $statusClass = "status-customer" }
                "Waiting on Customer" { $statusClass = "status-waiting" }
                "In Progress" { $statusClass = "status-inprogress" }
                "Quote/Billing" { $statusClass = "status-quotebilling" }
                default { $statusClass = "status-open" }
            }

            $statusColor = Get-StatusColor $statusName
            $bg = "#e6f2fb"
            $h.Add("<div class=`"status-col $statusClass`" style=`"border-color:$statusColor; background:$bg`">")
            $h.Add("<h4><span style=`"color:$statusColor`">$(Html-Encode($statusName))</span> <span class=`"badge`">$($sg.Count)</span></h4>")
            $h.Add("<div class=`"status-list`">")
            foreach ($t in ($sg.Group | Sort-Object AgeDays -Descending)) {
                $cust = if ($t.Customer) { $t.Customer } else { "Unknown Customer" }
                $prio = if ($t.Priority) { $t.Priority } else { "" }
                $prioTxt = if ($prio) { "($prio) " } else { "" }
                $ageTxt = if ($null -ne $t.AgeDays) { "$($t.AgeDays)d" } else { "Unknown age" }
                $badge = if ($t.LongOpen) { " <span class=`"badge badge-long`">LONG OPEN</span>" } else { "" }
                $line = "#$($t.Number) [$($t.Status)] $prioTxt$($t.Subject) — $cust (Age $ageTxt) | Tech: $Tech"
                $ticketUrl = Get-TicketUrl $subdomain $t.TicketId
                if ($ticketUrl) {
                    $h.Add("<a class=`"ticket-card`" style=`"border-left:4px solid $statusColor`" href=`"$(Html-Encode($ticketUrl))`" target=`"_blank`">$(Html-Encode($line))$badge</a>")
                } else {
                    $h.Add("<div class=`"ticket-card`" style=`"border-left:4px solid $statusColor`">$(Html-Encode($line))$badge</div>")
                }
            }
            $h.Add("</div>")
            $h.Add("</div>")
        }
        $h.Add("</div>")
    } else {
        $h.Add("<div class=`"muted`">No open tickets.</div>")
    }

    $h.Add("</div></body></html>")
    ($h -join "`r`n") | Out-File -FilePath $p -Encoding UTF8
}

function Write-TechClosedPage {
    param(
        [Parameter(Mandatory=$true)][string]$Tech,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Tickets,
        [Parameter(Mandatory=$true)][string]$FileName
    )

    $p = Join-Path $htmlSummaryDir $FileName
    $h = New-Object System.Collections.Generic.List[string]
    $h.Add("<!doctype html>")
    $h.Add("<html lang=`"en`">")
    $h.Add("<head>")
    $h.Add("  <meta charset=`"utf-8`" />")
    $h.Add("  <meta name=`"viewport`" content=`"width=device-width, initial-scale=1`" />")
    $h.Add("  <title>$(Html-Encode($Tech)) - Closed Tickets</title>")
    $h.Add("  <style>")
    $h.Add("    :root { --ink:#0f172a; --muted:#64748b; --panel:#ffffff; --outline:#e2e8f0; --accent:#0b4aa2; --accent-2:#0f6fb3; --bg1:#0b2a57; --bg2:#0b4aa2; --bg3:#0f6fb3; }")
    $h.Add("    body { margin:0; font-family:`"Manrope`",`"Avenir Next`",`"Trebuchet MS`",sans-serif; background:linear-gradient(140deg,var(--bg1),var(--bg2) 40%,var(--bg3) 85%); color:var(--ink); background-image:radial-gradient(rgba(255,255,255,0.06) 1px, transparent 1px); background-size:18px 18px; }")
    $h.Add("    .wrap { max-width:1200px; margin:0 auto; padding:28px 24px 40px; }")
    $h.Add("    .card { background:var(--panel); border-radius:16px; padding:16px 18px; box-shadow:0 10px 30px rgba(15,23,42,0.12); border:1px solid rgba(226,232,240,0.9); }")
    $h.Add("    .nav-btn { display:inline-flex; align-items:center; gap:8px; padding:8px 14px; border-radius:999px; background:rgba(255,255,255,0.95); border:1px solid rgba(226,232,240,0.9); color:#0f172a; text-decoration:none; font-weight:700; letter-spacing:0.02em; transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; }")
    $h.Add("    .nav-btn:hover { border-color:#94a3b8; box-shadow:0 8px 18px rgba(15,23,42,0.12); transform:translateY(-1px); }")
    $h.Add("    .muted { color:#e2e8f0; }")
    $h.Add("    .ticket-card { display:block; background:#ffffff; border:1px solid rgba(226,232,240,0.9); border-radius:12px; padding:10px 12px; margin:10px 0; color:inherit; text-decoration:none; box-shadow:0 6px 16px rgba(15,23,42,0.08); transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; }")
    $h.Add("    .ticket-card:hover { border-color:#94a3b8; box-shadow:0 10px 20px rgba(15,23,42,0.12); transform:translateY(-1px); }")
    $h.Add("  </style>")
    $h.Add("</head>")
    $h.Add("<body><div class=`"wrap`">")
    $techColor = Get-TechColor $Tech
    $h.Add("<div class=`"card`" style=`"background:#e6f2fb; color:#111827; border-color:#cfe2f3`">")
    $h.Add("<div><a class=`"nav-btn`" href=`"LatestReport.html`">Return to Summary</a></div>")
    $h.Add("<div style=`"background:$techColor; color:#ffffff; font-weight:700; padding:8px 10px; border-radius:8px; margin-top:8px;`">$(Html-Encode($Tech)) - Closed Tickets (Last Work Week)</div>")
    $h.Add("<div style=`"background:#0077c0; color:#ffffff; padding:6px 10px; border-radius:8px; margin-top:6px; display:inline-block;`">Week: $(Html-Encode($startUtc.ToString('yyyy-MM-dd'))) → $(Html-Encode($endUtc.AddSeconds(-1).ToString('yyyy-MM-dd')))</div>")
    $h.Add("<div style=`"background:#0077c0; color:#ffffff; padding:6px 10px; border-radius:8px; margin-top:6px; display:inline-block;`">Total closed: $(@($Tickets).Count)</div>")
    $h.Add("</div>")

    if (@($Tickets).Count -gt 0) {
        foreach ($t in ($Tickets | Sort-Object { Parse-Utc $_.UpdatedAt } -Descending)) {
            $cust = if ($t.Customer) { $t.Customer } else { "Unknown Customer" }
            $prio = if ($t.Priority) { $t.Priority } else { "" }
            $prioTxt = if ($prio) { "($prio) " } else { "" }
            $updated = if ($t.UpdatedAt) { (Parse-Utc $t.UpdatedAt).ToString('yyyy-MM-dd HH:mm') + " UTC" } else { "Unknown" }
            $line = "#$($t.Number) [Resolved] $prioTxt$($t.Subject) — $cust (Closed $updated) | Tech: $Tech"
            $ticketUrl = Get-TicketUrl $subdomain $t.TicketId
            if ($ticketUrl) {
                $h.Add("<a class=`"ticket-card`" style=`"border-left:4px solid $(Get-StatusColor 'Resolved')`" href=`"$(Html-Encode($ticketUrl))`" target=`"_blank`">$(Html-Encode($line))</a>")
            } else {
                $h.Add("<div class=`"ticket-card`" style=`"border-left:4px solid $(Get-StatusColor 'Resolved')`">$(Html-Encode($line))</div>")
            }
        }
    } else {
        $h.Add("<div class=`"muted`">No closed tickets.</div>")
    }

    $h.Add("</div></body></html>")
    ($h -join "`r`n") | Out-File -FilePath $p -Encoding UTF8
}

# -----------------------------
# HTML Report (collapsible)
# -----------------------------
$html = New-Object System.Collections.Generic.List[string]
$html.Add("<!doctype html>")
$html.Add("<html lang=`"en`">")
$html.Add("<head>")
$html.Add("  <meta charset=`"utf-8`" />")
$html.Add("  <meta name=`"viewport`" content=`"width=device-width, initial-scale=1`" />")
$html.Add("  <title>Syncro Tech Summary</title>")
$html.Add("  <style>")
    $html.Add("    :root { --ink:#0f172a; --muted:#64748b; --panel:#ffffff; --outline:#e2e8f0; --accent:#0b4aa2; --accent-2:#0f6fb3; --bg1:#071b35; --bg2:#0b4aa2; --bg3:#0f6fb3; }")
$html.Add("    body { margin:0; font-family:`"Manrope`",`"Avenir Next`",`"Trebuchet MS`",sans-serif; background:linear-gradient(140deg,var(--bg1),var(--bg2) 42%,var(--bg3) 90%); color:var(--ink); background-image:radial-gradient(rgba(255,255,255,0.06) 1px, transparent 1px); background-size:18px 18px; }")
$html.Add("    .wrap { max-width:1240px; margin:0 auto; padding:28px 24px 44px; }")
    $html.Add("    .header { position:relative; display:flex; flex-direction:column; align-items:flex-start; gap:10px; background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.18); border-radius:18px; padding:18px 18px 14px; box-shadow:0 18px 40px rgba(7,18,36,0.35); }")
    $html.Add("    .logo { height:68px; width:auto; }")
    $html.Add("    .logo-wrap { width:100%; }")
    $html.Add("    .logo-wrap .logo { display:block; margin-left:0; }")
    $html.Add("    h1 { margin:0; font-size:28px; letter-spacing:0.02em; }")
    $html.Add("    .meta { color:#e2e8f0; margin:14px 0 18px; text-align:left; display:inline-flex; gap:10px; align-items:center; background:rgba(255,255,255,0.14); border:1px solid rgba(255,255,255,0.2); padding:6px 12px; border-radius:999px; }")
$html.Add("    .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:16px; margin-bottom:24px; }")
    $html.Add("    .card { background:var(--panel); border-radius:16px; padding:14px 16px; box-shadow:0 10px 30px rgba(15,23,42,0.12); border:1px solid rgba(226,232,240,0.9); }")
    $html.Add("    .card-link { display:block; color:inherit; text-decoration:none; }")
    $html.Add("    .card-link:hover .card { border-color:#94a3b8; transform:translateY(-1px); }")
$html.Add("    .card .label { color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.12em; }")
$html.Add("    .card .value { font-size:22px; font-weight:700; margin-top:8px; }")
$html.Add("    .card.accent-blue { border-color:#bfdbfe; background:linear-gradient(135deg,#f8fafc,#eff6ff); }")
$html.Add("    .card.accent-green { border-color:#bbf7d0; background:linear-gradient(135deg,#f8fafc,#ecfccb); }")
$html.Add("    .card.accent-red { border-color:#fecaca; background:linear-gradient(135deg,#fff1f2,#fee2e2); }")
$html.Add("    .card.accent-orange { border-color:#fdba74; background:linear-gradient(135deg,#fff7ed,#ffedd5); }")
$html.Add("    .card.accent-yellow { border-color:#fde68a; background:linear-gradient(135deg,#fffbeb,#fef3c7); }")
$html.Add("    .summary-card summary { list-style:none; cursor:pointer; }")
$html.Add("    .summary-card summary::-webkit-details-marker { display:none; }")
$html.Add("    .card-list { margin-top:8px; max-height:240px; overflow:auto; padding-right:4px; }")
$html.Add("    .card-list ul { margin:6px 0 0 18px; }")
$html.Add("    .chips { display:flex; flex-wrap:wrap; gap:10px; margin-top:8px; width:100%; }")
$html.Add("    .chip { display:inline-flex; align-items:center; justify-content:space-between; gap:8px; padding:8px 12px; border-radius:999px; font-size:12px; font-weight:700; border:1px solid rgba(226,232,240,0.9); background:#f8fafc; flex:1 1 190px; }")
    $html.Add("    .chip { color:inherit; text-decoration:none; }")
    $html.Add("    .chip-blue { border-color:#bfdbfe; background:#eff6ff; }")
$html.Add("    .chip-green { border-color:#bbf7d0; background:#dcfce7; }")
$html.Add("    .chip-red { border-color:#fecaca; background:#fee2e2; }")
$html.Add("    .chip-orange { border-color:#fdba74; background:#ffedd5; }")
$html.Add("    .chip-yellow { border-color:#fde68a; background:#fef3c7; }")
$html.Add("    .chip-purple { border-color:#c4b5fd; background:#ede9fe; }")
$html.Add("    .chip-gray { border-color:#cbd5e1; background:#f1f5f9; }")
$html.Add("    details { background:var(--panel); border-radius:14px; padding:12px 14px; margin-bottom:12px; box-shadow:0 8px 18px rgba(15,23,42,0.1); border:1px solid rgba(226,232,240,0.9); }")
$html.Add("    summary { cursor:pointer; font-weight:700; letter-spacing:0.02em; }")
    $html.Add("    .header h1 { color:#ffffff; font-weight:700; }")
    $html.Add("    .section-title { margin:22px 0 10px 0; font-size:13px; color:#e2e8f0; font-weight:700; letter-spacing:0.18em; text-transform:uppercase; }")
$html.Add("    ul { margin:8px 0 0 20px; }")
    $html.Add("    .muted { color:var(--muted); }")
    $html.Add("    .badge { display:inline-block; padding:2px 8px; border-radius:999px; background:#e2e8f0; font-size:12px; margin-left:6px; }")
    $html.Add("    .badge-long { background:#fecaca; border:1px solid #fca5a5; color:#7f1d1d; }")
    $html.Add("    .bar-row { display:flex; align-items:center; gap:10px; margin:6px 0; }")
    $html.Add("    .bar-label { min-width:140px; font-size:12px; color:var(--muted); }")
    $html.Add("    .bar-track { flex:1; background:#e5e7eb; border-radius:999px; height:10px; overflow:hidden; }")
    $html.Add("    .bar-fill { height:100%; background:var(--accent); }")
    $html.Add("    .bar-val { width:40px; text-align:right; font-size:12px; color:var(--muted); }")
    $html.Add("    .bar-link { display:flex; align-items:center; gap:10px; text-decoration:none; color:inherit; padding:8px 10px; border-radius:12px; transition:background .15s ease; }")
    $html.Add("    .bar-link:hover { background:rgba(255,255,255,0.65); }")
    $html.Add("    .tech-links { display:flex; flex-wrap:wrap; gap:8px; margin:8px 0 12px 0; }")
    $html.Add("    .tech-link { display:inline-block; padding:8px 12px; border-radius:999px; background:#f1f5f9; border:1px solid #e2e8f0; font-size:12px; text-decoration:none; color:inherit; font-weight:700; }")
$html.Add("    .status-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:12px; }")
$html.Add("    .status-col { background:rgba(255,255,255,0.95); border:1px solid rgba(226,232,240,0.9); border-radius:14px; padding:12px; box-shadow:0 8px 18px rgba(15,23,42,0.08); }")
$html.Add("    .status-col h4 { margin:0 0 10px 0; font-size:14px; letter-spacing:0.02em; }")
$html.Add("    .status-col ul { margin:0 0 0 18px; }")
$html.Add("    .status-list { max-height:420px; overflow:auto; padding-right:4px; }")
$html.Add("    .status-open { border-color:#f59e0b; background:#fef3c7; }")
$html.Add("    .status-waiting { border-color:#f97316; background:#ffedd5; }")
$html.Add("    .status-customer { border-color:#ef4444; background:#fee2e2; }")
$html.Add("    .status-resolved { border-color:#22c55e; background:#dcfce7; }")
$html.Add("    .status-inprogress { border-color:#3b82f6; background:#dbeafe; }")
$html.Add("    .status-quotebilling { border-color:#8b5cf6; background:#ede9fe; }")
$html.Add("    .ticket-card { background:#ffffff; border:1px solid rgba(226,232,240,0.9); border-radius:12px; padding:10px 12px; margin:10px 0; box-shadow:0 6px 16px rgba(15,23,42,0.08); transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; }")
$html.Add("    .ticket-card.status-customer { background:#fecaca; }")
$html.Add("    .ticket-card.status-waiting { background:#fed7aa; }")
$html.Add("    .ticket-card.status-inprogress { background:#bfdbfe; }")
$html.Add("    .ticket-card.status-quotebilling { background:#ddd6fe; }")
$html.Add("    .ticket-card.status-resolved { background:#bbf7d0; }")
$html.Add("    .ticket-card.status-open { background:#fde68a; }")
$html.Add("    .ticket-card summary { list-style:none; cursor:pointer; }")
$html.Add("    .ticket-card summary::-webkit-details-marker { display:none; }")
$html.Add("    .ticket-title { font-weight:700; }")
$html.Add("    .ticket-sub { color:var(--muted); font-size:12px; }")
$html.Add("    .ticket-card:hover { border-color:#94a3b8; box-shadow:0 10px 20px rgba(15,23,42,0.12); transform:translateY(-1px); }")
$html.Add("    .ticket-body { margin-top:6px; font-size:13px; }")
$html.Add("  </style>")
$html.Add("</head>")
$html.Add("<body>")
$html.Add("<div class=`"wrap`">")
$html.Add("<div class=`"header`">")
if (Test-Path $logoTarget) {
    $html.Add("<div class=`"logo-wrap`"><img class=`"logo`" src=`"logo.png`" alt=`"Bigfoot Networks logo`" /></div>")
}
$html.Add("<h1>$(Html-Encode($reportTitle))</h1>")
$html.Add("</div>")
if ($windowLocalStart -and $windowLocalEnd) {
    $html.Add("<div class=`"meta`">Week: $(Html-Encode($windowLocalStart.ToString('yyyy-MM-dd'))) → $(Html-Encode($windowLocalEnd.AddSeconds(-1).ToString('yyyy-MM-dd')))</div>")
} else {
    $html.Add("<div class=`"meta`">Week: $(Html-Encode($startUtc.ToString('yyyy-MM-dd'))) → $(Html-Encode($endUtc.AddSeconds(-1).ToString('yyyy-MM-dd')))</div>")
}

$html.Add("<div class=`"section-title`">Last Week</div>")
$html.Add("<div class=`"cards`">")
$html.Add("<a class=`"card-link`" href=`"Summary_TicketsUpdated.html`"><div class=`"card accent-blue`"><div class=`"label`">Tickets Updated</div><div class=`"value`">$(@($ticketRows).Count)</div></div></a>")
$html.Add("<a class=`"card-link`" href=`"Summary_Resolved.html`"><div class=`"card accent-green`"><div class=`"label`">Closed (Resolved)</div><div class=`"value`">$totalClosed</div></div></a>")
$html.Add("<a class=`"card-link`" href=`"Summary_Created.html`"><div class=`"card accent-blue`"><div class=`"label`">Tickets Opened</div><div class=`"value`">$(@($ticketsCreated).Count)</div></div></a>")
$html.Add("<a class=`"card-link`" href=`"Summary_Open.html`"><div class=`"card`"><div class=`"label`">Opened Last Week (Still Open)</div><div class=`"value`">$(@($openLastWeekUnresolved).Count)</div></div></a>")
$html.Add("</div>")

$html.Add("<div class=`"section-title`">Current</div>")
$html.Add("<div class=`"cards`">")
$html.Add("<a class=`"card-link`" href=`"Summary_StaleNoUpdate.html`"><div class=`"card accent-yellow`"><div class=`"label`">Stale (No Update)</div><div class=`"value`">$(@($staleNoUpdate).Count)</div></div></a>")
$html.Add("<a class=`"card-link`" href=`"Summary_CustomerReply.html`"><div class=`"card accent-red`"><div class=`"label`">Customer Reply &gt; $staleCustomerReplyHours h</div><div class=`"value`">$(@($staleCustomerReply).Count)</div></div></a>")
$html.Add("<a class=`"card-link`" href=`"Summary_WaitingOnCustomer.html`"><div class=`"card accent-orange`"><div class=`"label`">Waiting on Customer &gt; $staleWaitingOnCustomerDays d</div><div class=`"value`">$(@($staleWaitingOnCustomer).Count)</div></div></a>")
$html.Add("<a class=`"card-link`" href=`"Summary_LongOpen.html`"><div class=`"card accent-red`"><div class=`"label`">Long Open &gt;= $longOpenDays d</div><div class=`"value`">$totalLongOpen</div></div></a>")
$html.Add("</div>")


# Write summary pages
$ticketsUpdated = @($ticketRows | Sort-Object { Parse-Utc $_.UpdatedAt } -Descending)
$ticketsResolved = @($ticketRows | Where-Object { $_.Status -eq "Resolved" } | Sort-Object { Parse-Utc $_.UpdatedAt } -Descending)
$ticketsStaleNoUpdate = @($staleNoUpdate | Sort-Object { Parse-Utc $_.UpdatedAt } -Descending)
$ticketsCustomerReply = @($staleCustomerReply | Sort-Object { Parse-Utc $_.UpdatedAt } -Descending)
$ticketsWaitingOnCustomer = @($staleWaitingOnCustomer | Sort-Object { Parse-Utc $_.UpdatedAt } -Descending)

Write-SummaryPage -Title "Tickets Updated" -FileName "Summary_TicketsUpdated.html" -Tickets $ticketsUpdated -Accent "accent-blue"
Write-SummaryPage -Title "Closed (Resolved)" -FileName "Summary_Resolved.html" -Tickets $ticketsResolved -Accent "accent-green"
Write-SummaryPage -Title "Tickets Opened" -FileName "Summary_Created.html" -Tickets $ticketsCreated -Accent "accent-blue"
Write-SummaryPage -Title "Stale (No Update)" -FileName "Summary_StaleNoUpdate.html" -Tickets $ticketsStaleNoUpdate -Accent "accent-yellow"
Write-SummaryPage -Title "Customer Reply > $staleCustomerReplyHours h" -FileName "Summary_CustomerReply.html" -Tickets $ticketsCustomerReply -Accent "accent-red"
Write-SummaryPage -Title "Waiting on Customer > $staleWaitingOnCustomerDays d" -FileName "Summary_WaitingOnCustomer.html" -Tickets $ticketsWaitingOnCustomer -Accent "accent-orange"
Write-SummaryPage -Title "Opened Last Week (Still Open)" -FileName "Summary_Open.html" -Tickets $openLastWeekUnresolved -Accent "accent-blue"
$longOpenTickets = $openTicketRows | Where-Object { $_.LongOpen }
Write-SummaryPage -Title "Long Open (>= $longOpenDays d)" -FileName "Summary_LongOpen.html" -Tickets $longOpenTickets -Accent "accent-red"

# Per-tech open ticket pages
$techNamesForPages =
    @(
        $byTechTickets | ForEach-Object { $_.Name }
        $openByTechMap.Keys
    ) | Where-Object { $_ } | Select-Object -Unique | Sort-Object
foreach ($t in $techNamesForPages) {
    $fileName = "Open_" + (Slugify $t) + ".html"
    $tickets = @()
    if ($openByTechMap.ContainsKey($t) -and $null -ne $openByTechMap[$t]) {
        $tickets = @($openByTechMap[$t])
    }
    Write-TechOpenPage -Tech $t -Tickets $tickets -FileName $fileName
}

# Per-tech closed ticket pages (last work week)
$techNamesClosedPages =
    @(
        $byTechTickets | ForEach-Object { $_.Name }
    ) | Where-Object { $_ } | Select-Object -Unique | Sort-Object
foreach ($t in $techNamesClosedPages) {
    $fileName = "Closed_" + (Slugify $t) + ".html"
    $grp = $byTechTickets | Where-Object { $_.Name -eq $t } | Select-Object -First 1
    $tickets = if ($grp) { @($grp.Group | Where-Object { $_.Status -eq "Resolved" }) } else { @() }
    Write-TechClosedPage -Tech $t -Tickets $tickets -FileName $fileName
}

$html.Add("<div class=`"section-title`">Per Technician (Open Tickets)</div>")
$html.Add("<div class=`"card`">")
$techNames =
    @(
        $byTechTickets | ForEach-Object { $_.Name }
        $openByTechMap.Keys
    ) | Where-Object { $_ } | Select-Object -Unique | Sort-Object
$maxOpen = 0
foreach ($t in $techNames) {
    $count = if ($openByTechMap.ContainsKey($t)) { @($openByTechMap[$t]).Count } else { 0 }
    if ($count -gt $maxOpen) { $maxOpen = $count }
}
$html.Add("<div class=`"tech-links`">")
foreach ($t in $techNames) {
    $fileName = "Open_" + (Slugify $t) + ".html"
    $tColor = Get-TechColor $t
    $html.Add("<a class=`"tech-link`" style=`"border-color:$tColor; background:$tColor; color:#fff`" href=`"$fileName`">$(Html-Encode($t))</a>")
}
$html.Add("</div>")
foreach ($t in $techNames) {
    $count = if ($openByTechMap.ContainsKey($t)) { @($openByTechMap[$t]).Count } else { 0 }
    $pct = if ($maxOpen -gt 0) { [math]::Round(($count / $maxOpen) * 100, 0) } else { 0 }
    $fileName = "Open_" + (Slugify $t) + ".html"
    $tColor = Get-TechColor $t
    $html.Add("<a class=`"bar-link`" href=`"$fileName`">")
    $html.Add("<div class=`"bar-label`" style=`"color:$tColor; font-weight:700`">$(Html-Encode($t))</div>")
    $html.Add("<div class=`"bar-track`"><div class=`"bar-fill`" style=`"width:$pct%; background:$tColor`"></div></div>")
    $html.Add("<div class=`"bar-val`">$count</div>")
    $html.Add("</a>")
}
$html.Add("</div>")

$html.Add("<div class=`"section-title`">Per Technician (Closed Tickets - Last Work Week)</div>")
$html.Add("<div class=`"card`">")
$techNamesClosed =
    @(
        $byTechTickets | ForEach-Object { $_.Name }
    ) | Where-Object { $_ } | Select-Object -Unique | Sort-Object
$maxClosed = 0
foreach ($t in $techNamesClosed) {
    $grp = $byTechTickets | Where-Object { $_.Name -eq $t } | Select-Object -First 1
    $count = if ($grp) { @($grp.Group | Where-Object { $_.Status -eq "Resolved" }).Count } else { 0 }
    if ($count -gt $maxClosed) { $maxClosed = $count }
}
$html.Add("<div class=`"tech-links`">")
foreach ($t in $techNamesClosed) {
    $fileName = "Closed_" + (Slugify $t) + ".html"
    $tColor = Get-TechColor $t
    $html.Add("<a class=`"tech-link`" style=`"border-color:$tColor; background:$tColor; color:#fff`" href=`"$fileName`">$(Html-Encode($t))</a>")
}
$html.Add("</div>")
foreach ($t in $techNamesClosed) {
    $grp = $byTechTickets | Where-Object { $_.Name -eq $t } | Select-Object -First 1
    $count = if ($grp) { @($grp.Group | Where-Object { $_.Status -eq "Resolved" }).Count } else { 0 }
    $pct = if ($maxClosed -gt 0) { [math]::Round(($count / $maxClosed) * 100, 0) } else { 0 }
    $fileName = "Closed_" + (Slugify $t) + ".html"
    $tColor = Get-TechColor $t
    $html.Add("<a class=`"bar-link`" href=`"$fileName`">")
    $html.Add("<div class=`"bar-label`" style=`"color:$tColor; font-weight:700`">$(Html-Encode($t))</div>")
    $html.Add("<div class=`"bar-track`"><div class=`"bar-fill`" style=`"width:$pct%; background:$tColor`"></div></div>")
    $html.Add("<div class=`"bar-val`">$count</div>")
    $html.Add("</a>")
}
$html.Add("</div>")
$html.Add("</div></body></html>")
($html -join "`r`n") | Out-File -FilePath $htmlReportPath -Encoding UTF8
Write-Host "HTML report written to: $htmlReportPath"

Stop-Transcript | Out-Null




