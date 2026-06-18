param(
  [string]$Symbol = "NIFTY",
  [string]$LocalDatPath = (Join-Path $PSScriptRoot "data\nifty-options.dat"),
  [string]$RepoOwner = "pjalpesh1",
  [string]$RepoName = "nifty-option-analyse",
  [string]$RepoPath = "data/nifty-options.dat",
  [string]$Branch = "main",
  [int]$IntervalSeconds = 60,
  [int]$ExpiryCount = 2,
  [string]$GitHubToken = $env:GITHUB_TOKEN,
  [switch]$NoUpload,
  [switch]$Once,
  [switch]$RunOnlyMarketHours
)

$ErrorActionPreference = "Stop"
$culture = [System.Globalization.CultureInfo]::InvariantCulture
$NseBaseUrl = "https://www.nseindia.com"
$OptionChainPage = "$NseBaseUrl/option-chain"
$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"

Add-Type -AssemblyName System.Web.Extensions
$jsonParser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$jsonParser.MaxJsonLength = 2147483647

function Get-IndiaNow {
  try {
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("India Standard Time")
    return [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
  } catch {
    return (Get-Date)
  }
}

function Test-IndianMarketOpen {
  $now = Get-IndiaNow
  if ($now.DayOfWeek -eq "Saturday" -or $now.DayOfWeek -eq "Sunday") { return $false }
  $open = $now.Date.AddHours(9).AddMinutes(15)
  $close = $now.Date.AddHours(15).AddMinutes(30)
  return ($now -ge $open -and $now -le $close)
}

function New-NseSession {
  $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
  $headers = @{
    "User-Agent" = $UserAgent
    "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    "Accept-Language" = "en-US,en;q=0.9"
    "Cache-Control" = "no-cache"
  }
  Invoke-WebRequest -Uri $OptionChainPage -Headers $headers -WebSession $session -UseBasicParsing -TimeoutSec 30 | Out-Null
  return $session
}

function Invoke-NseJson {
  param(
    [Microsoft.PowerShell.Commands.WebRequestSession]$Session,
    [string]$Path
  )

  $headers = @{
    "User-Agent" = $UserAgent
    "Accept" = "application/json,text/plain,*/*"
    "Accept-Language" = "en-US,en;q=0.9"
    "Referer" = $OptionChainPage
    "Cache-Control" = "no-cache"
  }
  $uri = if ($Path.StartsWith("http")) { $Path } else { "$NseBaseUrl$Path" }
  $response = Invoke-WebRequest -Uri $uri -Headers $headers -WebSession $Session -UseBasicParsing -TimeoutSec 30
  $text = [string]$response.Content
  if ([string]::IsNullOrWhiteSpace($text)) { throw "NSE returned an empty response." }
  return $jsonParser.DeserializeObject($text)
}

function Get-DictValue {
  param(
    [object]$Dict,
    [string[]]$Names
  )

  if ($null -eq $Dict) { return $null }
  foreach ($name in $Names) {
    try {
      if ($Dict.ContainsKey($name) -and $null -ne $Dict[$name]) {
        return $Dict[$name]
      }
    } catch {}
  }
  return $null
}

function Convert-NseDate {
  param([object]$Value)

  if ($null -eq $Value) { return "" }
  $text = ([string]$Value).Trim()
  if (!$text) { return "" }
  if ($text -match '^(\d{1,2})-(\d{1,2})-(\d{4})$') {
    return "{0}-{1}-{2}" -f $matches[3], $matches[2].PadLeft(2, "0"), $matches[1].PadLeft(2, "0")
  }
  if ($text -match '^(\d{1,2})-([A-Za-z]{3,})-(\d{4})$') {
    try {
      return ([DateTime]::ParseExact($text, "dd-MMM-yyyy", $culture)).ToString("yyyy-MM-dd", $culture)
    } catch {}
  }
  return $text
}

function Convert-ToDatValue {
  param([object]$Value)

  if ($null -eq $Value) { return "" }
  if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
    return ([System.Convert]::ToString($Value, $culture))
  }
  return (([string]$Value) -replace "[\r\n|]+", " ").Trim()
}

function Export-NseChainToDat {
  param(
    [object]$ContractInfo,
    [Microsoft.PowerShell.Commands.WebRequestSession]$Session,
    [string]$DatPath
  )

  $expiryDates = Get-DictValue $ContractInfo @("expiryDates")
  if (!$expiryDates -or $expiryDates.Count -eq 0) {
    throw "NSE contract-info returned no expiry dates for $Symbol."
  }

  $selectedExpiries = @($expiryDates | Select-Object -First ([Math]::Max(1, $ExpiryCount)))
  $headers = @(
    "expiryDate",
    "strikePrice",
    "optionType",
    "openInterest",
    "changeinOpenInterest",
    "lastPrice",
    "change",
    "impliedVolatility",
    "delta",
    "gamma",
    "theta",
    "totalTradedVolume"
  )

  $lines = New-Object System.Collections.Generic.List[string]
  $lines.Add(($headers -join "|"))
  $rowCount = 0
  $spot = 0
  $timestamp = ""

  foreach ($expiry in $selectedExpiries) {
    $expiryText = [string]$expiry
    $path = "/api/option-chain-v3?type=Indices&symbol=$([uri]::EscapeDataString($Symbol))&expiry=$([uri]::EscapeDataString($expiryText))"
    $chain = Invoke-NseJson -Session $Session -Path $path
    if (!$chain -or $chain.Count -eq 0) {
      throw "NSE returned blank option-chain data for $Symbol $expiryText."
    }

    $records = Get-DictValue $chain @("records")
    $data = Get-DictValue $records @("data")
    if ($spot -eq 0) {
      $rawSpot = Get-DictValue $records @("underlyingValue")
      if ($rawSpot) { $spot = [double]$rawSpot }
    }
    if (!$timestamp) { $timestamp = [string](Get-DictValue $records @("timestamp")) }

    foreach ($item in @($data)) {
      if ($null -eq $item) { continue }
      foreach ($type in @("CE", "PE")) {
        $leg = Get-DictValue $item @($type)
        if ($null -eq $leg) { continue }

        $expiryDate = Convert-NseDate (Get-DictValue $leg @("expiryDate", "expiryDates"))
        if (!$expiryDate) { $expiryDate = Convert-NseDate $expiryText }
        $strikePrice = Get-DictValue $leg @("strikePrice")
        if ($null -eq $strikePrice) { $strikePrice = Get-DictValue $item @("strikePrice") }

        $fields = @(
          $expiryDate,
          $strikePrice,
          $type,
          (Get-DictValue $leg @("openInterest")),
          (Get-DictValue $leg @("changeinOpenInterest", "changeInOpenInterest")),
          (Get-DictValue $leg @("lastPrice")),
          (Get-DictValue $leg @("change")),
          (Get-DictValue $leg @("impliedVolatility")),
          (Get-DictValue $leg @("delta")),
          (Get-DictValue $leg @("gamma")),
          (Get-DictValue $leg @("theta")),
          (Get-DictValue $leg @("totalTradedVolume"))
        )
        $lines.Add((($fields | ForEach-Object { Convert-ToDatValue $_ }) -join "|"))
        $rowCount += 1
      }
    }
  }

  if ($rowCount -eq 0) {
    throw "No option rows found from NSE. Website upload skipped."
  }

  $datDir = Split-Path -Parent $DatPath
  if ($datDir -and !(Test-Path -LiteralPath $datDir)) {
    New-Item -ItemType Directory -Path $datDir | Out-Null
  }
  [System.IO.File]::WriteAllLines($DatPath, $lines, [System.Text.UTF8Encoding]::new($false))

  return @{
    Rows = $rowCount
    Spot = $spot
    Timestamp = $timestamp
    Expiries = ($selectedExpiries -join ", ")
  }
}

function Get-Token {
  if ($GitHubToken) { return $GitHubToken }

  $gh = Get-Command gh -ErrorAction SilentlyContinue
  if (!$gh -and (Test-Path "C:\Program Files\GitHub CLI\gh.exe")) {
    $gh = Get-Item "C:\Program Files\GitHub CLI\gh.exe"
  }

  if ($gh) {
    $token = & $gh.Source auth token 2>$null
    if ($LASTEXITCODE -eq 0 -and $token) { return $token.Trim() }
  }

  return ""
}

function Invoke-GitHubApi {
  param(
    [string]$Method,
    [string]$Path,
    [object]$Body,
    [string]$Token
  )

  $headers = @{
    Accept = "application/vnd.github+json"
    Authorization = "Bearer $Token"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent" = "nifty-nse-site-uploader"
  }

  $uri = "https://api.github.com$Path"
  if ($null -eq $Body) {
    return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
  }

  return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -Body ($Body | ConvertTo-Json -Depth 8) -ContentType "application/json"
}

function Upload-DatToGitHub {
  param(
    [string]$DatPath,
    [string]$Token
  )

  if (!$Token) {
    Write-Warning "No GitHub token found. DAT saved locally only. Run gh auth login or set GITHUB_TOKEN."
    return $false
  }

  $endpoint = "/repos/$RepoOwner/$RepoName/contents/$RepoPath"
  $existing = $null
  try {
    $existing = Invoke-GitHubApi -Method "GET" -Path $endpoint -Token $Token
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 404) { throw }
  }

  $body = @{
    message = "Update NSE NIFTY DAT feed"
    content = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($DatPath))
    branch = $Branch
  }
  if ($existing -and $existing.sha) { $body.sha = $existing.sha }

  Invoke-GitHubApi -Method "PUT" -Path $endpoint -Body $body -Token $Token | Out-Null
  return $true
}

$lastHash = ""
$session = $null

Write-Host "NSE site DAT updater started. Press Ctrl+C to stop."
Write-Host "Source: $OptionChainPage"
Write-Host "Symbol: $Symbol"
Write-Host "DAT:    $LocalDatPath"
Write-Host "Web:    https://$RepoOwner.github.io/$RepoName/$RepoPath"
Write-Host "Note: NSE free site data can be delayed/blocked by NSE. This script never uploads blank data."

do {
  try {
    if ($RunOnlyMarketHours -and !(Test-IndianMarketOpen)) {
      Write-Host ("Indian market closed at {0}. Waiting..." -f (Get-IndiaNow).ToString("dd-MMM-yyyy HH:mm:ss"))
      if ($Once) { break }
      Start-Sleep -Seconds $IntervalSeconds
      continue
    }

    if (!$session) { $session = New-NseSession }

    $contractInfo = Invoke-NseJson -Session $session -Path "/api/option-chain-contract-info?symbol=$([uri]::EscapeDataString($Symbol))"
    $result = Export-NseChainToDat -ContractInfo $contractInfo -Session $session -DatPath $LocalDatPath
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $LocalDatPath).Hash

    if ($hash -ne $lastHash) {
      if ($NoUpload) {
        $uploaded = $false
      } else {
        $token = Get-Token
        $uploaded = Upload-DatToGitHub -DatPath $LocalDatPath -Token $token
      }
      $lastHash = $hash
      $status = if ($uploaded) { "uploaded" } else { "local only" }
      Write-Host ("{0} rows {1} at {2}; spot {3}; expiries {4}; NSE time {5}" -f $result.Rows, $status, (Get-IndiaNow).ToString("HH:mm:ss"), $result.Spot, $result.Expiries, $result.Timestamp)
    } else {
      Write-Host ("No data change at {0}; spot {1}; NSE time {2}" -f (Get-IndiaNow).ToString("HH:mm:ss"), $result.Spot, $result.Timestamp)
    }
  } catch {
    Write-Warning $_.Exception.Message
    Write-Warning "If this repeats, open https://www.nseindia.com/option-chain once in your browser, then restart this BAT file."
    $session = $null
  }

  if ($Once) { break }
  Start-Sleep -Seconds $IntervalSeconds
} while ($true)
