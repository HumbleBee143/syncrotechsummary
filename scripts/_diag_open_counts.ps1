$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = Split-Path -Parent $scriptDir
$cfgCandidates = @(
    (Join-Path $projectRoot "config\Syncro-TechSummary.config.json"),
    (Join-Path $scriptDir "Syncro-TechSummary.config.json"),
    (Join-Path $projectRoot "Syncro-TechSummary.config.json")
)
$cfgPath = $cfgCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($cfgPath)) { throw "Config not found. Checked: $($cfgCandidates -join ', ')" }
$cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
$sub = $cfg.Subdomain
$key = $cfg.ApiKey
$openStatuses = @($cfg.OpenTickets.Statuses)

function Has-Prop($o,$n){($null -ne $o) -and ($null -ne $o.PSObject.Properties[$n])}
function Get-Prop($o,$n){if(Has-Prop $o $n){$o.PSObject.Properties[$n].Value}else{$null}}
function Get-SyncroPaged([string]$UrlBase,[string]$ItemsKey){
    $all=New-Object System.Collections.Generic.List[object]
    $page=1
    $total=1
    do{
        $url = if($UrlBase -match '\?'){"$UrlBase&page=$page"}else{"$UrlBase?page=$page"}
        $resp=Invoke-RestMethod -Method GET -Uri $url
        $items=Get-Prop $resp $ItemsKey
        if($items){foreach($i in $items){$all.Add($i)}}
        $meta=Get-Prop $resp 'meta'
        $total=1
        if($meta -and (Has-Prop $meta 'total_pages') -and $meta.total_pages){$total=[int]$meta.total_pages}
        $page++
    } while($page -le $total)
    return $all
}
function Get-Tech($t){$u=Get-Prop $t 'user'; if($u -and (Has-Prop $u 'full_name') -and $u.full_name){[string]$u.full_name}else{"Unassigned"}}

$raw=New-Object System.Collections.Generic.List[object]
foreach($st in $openStatuses){
    if([string]::IsNullOrWhiteSpace($st)){continue}
    $stEnc=[uri]::EscapeDataString($st)
    $url="https://$sub.syncromsp.com/api/v1/tickets?api_key=$key&status=$stEnc"
    $rows=Get-SyncroPaged -UrlBase $url -ItemsKey 'tickets'
    if($rows){foreach($r in $rows){$raw.Add($r)}}
}
$byId=@{}
foreach($t in $raw){$id=Get-Prop $t 'id'; if($id -and -not $byId.ContainsKey($id)){$byId[$id]=$t}}
$dedup=$byId.Values
$counts=$dedup | Group-Object { Get-Tech $_ } | Sort-Object Name | Select-Object Name,Count
$counts | Format-Table -AutoSize
