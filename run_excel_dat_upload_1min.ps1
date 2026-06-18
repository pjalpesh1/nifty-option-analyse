param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot "nifty_option_data_cache.xlsx"),
  [string]$SheetName = "OptionData",
  [string]$LocalDatPath = (Join-Path $PSScriptRoot "data\nifty-options.dat"),
  [string]$RepoOwner = "pjalpesh1",
  [string]$RepoName = "nifty-option-analyse",
  [string]$RepoPath = "data/nifty-options.dat",
  [string]$Branch = "main",
  [int]$IntervalSeconds = 60,
  [string]$GitHubToken = $env:GITHUB_TOKEN,
  [switch]$NoUpload,
  [switch]$AllowEmptyUpload,
  [switch]$Once
)

$ErrorActionPreference = "Stop"
$culture = [System.Globalization.CultureInfo]::InvariantCulture

function Convert-ToDatValue {
  param(
    [object]$Value,
    [string]$Header
  )

  if ($null -eq $Value) { return "" }

  if ($Header -match "date|expiry") {
    if ($Value -is [datetime]) {
      return $Value.ToString("dd-MMM-yyyy", $culture)
    }
    if ($Value -is [double] -or $Value -is [int]) {
      try {
        return ([datetime]::FromOADate([double]$Value)).ToString("dd-MMM-yyyy", $culture)
      } catch {}
    }
  }

  $text = [System.Convert]::ToString($Value, $culture)
  return ($text -replace "[\r\n|]+", " ").Trim()
}

function Export-ExcelToDat {
  param(
    [string]$WorkbookPath,
    [string]$WorksheetName,
    [string]$DatPath
  )

  if (!(Test-Path -LiteralPath $WorkbookPath)) {
    throw "Excel file not found: $WorkbookPath"
  }

  $fullWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
  $datDir = Split-Path -Parent $DatPath
  if ($datDir -and !(Test-Path -LiteralPath $datDir)) {
    New-Item -ItemType Directory -Path $datDir | Out-Null
  }

  $excel = $null
  $workbook = $null
  $sheet = $null

  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($fullWorkbookPath, $null, $true)
    $sheet = $workbook.Worksheets.Item($WorksheetName)
    $used = $sheet.UsedRange
    $lastRow = [Math]::Max(1, $used.Row + $used.Rows.Count - 1)
    $lastCol = [Math]::Max(1, $used.Column + $used.Columns.Count - 1)

    $headers = New-Object System.Collections.Generic.List[string]
    for ($col = 1; $col -le $lastCol; $col++) {
      $header = Convert-ToDatValue $sheet.Cells.Item(1, $col).Value2 ""
      if ([string]::IsNullOrWhiteSpace($header)) { break }
      $headers.Add($header)
    }

    if ($headers.Count -lt 3) {
      throw "Header row missing in $WorksheetName. Keep row 1 headers unchanged."
    }

    $expiryIndex = [Array]::IndexOf($headers.ToArray(), "expiryDate")
    $strikeIndex = [Array]::IndexOf($headers.ToArray(), "strikePrice")
    $typeIndex = [Array]::IndexOf($headers.ToArray(), "optionType")
    if ($expiryIndex -lt 0 -or $strikeIndex -lt 0 -or $typeIndex -lt 0) {
      throw "Required headers missing. Need expiryDate, strikePrice, optionType."
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(($headers -join "|"))
    $rowCount = 0

    for ($row = 2; $row -le $lastRow; $row++) {
      $fields = New-Object System.Collections.Generic.List[string]
      for ($col = 1; $col -le $headers.Count; $col++) {
        $fields.Add((Convert-ToDatValue $sheet.Cells.Item($row, $col).Value2 $headers[$col - 1]))
      }

      if ([string]::IsNullOrWhiteSpace($fields[$expiryIndex]) -or
          [string]::IsNullOrWhiteSpace($fields[$strikeIndex]) -or
          [string]::IsNullOrWhiteSpace($fields[$typeIndex])) {
        continue
      }

      $lines.Add(($fields -join "|"))
      $rowCount += 1
    }

    [System.IO.File]::WriteAllLines($DatPath, $lines, [System.Text.UTF8Encoding]::new($false))
    return $rowCount
  } finally {
    if ($workbook) { $workbook.Close($false) | Out-Null }
    if ($excel) { $excel.Quit() | Out-Null }
    foreach ($obj in @($sheet, $workbook, $excel)) {
      if ($obj) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

function Get-Token {
  if ($GitHubToken) { return $GitHubToken }

  $gh = Get-Command gh -ErrorAction SilentlyContinue
  if (!$gh -and (Test-Path "C:\Program Files\GitHub CLI\gh.exe")) {
    $gh = Get-Item "C:\Program Files\GitHub CLI\gh.exe"
  }

  if ($gh) {
    $ghPath = if ($gh.Source) { $gh.Source } else { $gh.FullName }
    $token = & $ghPath auth token 2>$null
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
    "User-Agent" = "nifty-excel-dat-uploader"
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
    Write-Warning "No GitHub token found. DAT exported locally only. Set GITHUB_TOKEN or run gh auth login."
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
    message = "Update NIFTY DAT feed"
    content = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($DatPath))
    branch = $Branch
  }
  if ($existing -and $existing.sha) { $body.sha = $existing.sha }

  Invoke-GitHubApi -Method "PUT" -Path $endpoint -Body $body -Token $Token | Out-Null
  return $true
}

$lastHash = ""
Write-Host "Excel DAT uploader started. Press Ctrl+C to stop."
Write-Host "Excel: $ExcelPath"
Write-Host "DAT:   $LocalDatPath"
Write-Host "Web:   https://$RepoOwner.github.io/$RepoName/$RepoPath"

do {
  try {
    $rows = Export-ExcelToDat -WorkbookPath $ExcelPath -WorksheetName $SheetName -DatPath $LocalDatPath
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $LocalDatPath).Hash

    if ($rows -eq 0 -and !$AllowEmptyUpload) {
      Write-Warning "No option rows found in Excel sheet OptionData. Upload skipped so website keeps previous data."
      Write-Warning "Paste option-chain rows under row 1 headers, save Excel, then wait for next cycle."
      if ($hash -ne $lastHash) { $lastHash = $hash }
      if ($Once) { break }
      Start-Sleep -Seconds $IntervalSeconds
      continue
    }

    if ($hash -ne $lastHash) {
      if ($NoUpload) {
        $uploaded = $false
      } else {
        $token = Get-Token
        $uploaded = Upload-DatToGitHub -DatPath $LocalDatPath -Token $token
      }
      $lastHash = $hash
      $status = if ($uploaded) { "uploaded" } else { "local only" }
      Write-Host ("{0} rows exported, {1} at {2}" -f $rows, $status, (Get-Date -Format "HH:mm:ss"))
      if ($rows -eq 0) {
        Write-Warning "No option rows found. Add data in Excel sheet OptionData and save the file."
      }
    } else {
      Write-Host ("No data change at {0}" -f (Get-Date -Format "HH:mm:ss"))
    }
  } catch {
    Write-Warning $_.Exception.Message
  }

  if ($Once) { break }
  Start-Sleep -Seconds $IntervalSeconds
} while ($true)
