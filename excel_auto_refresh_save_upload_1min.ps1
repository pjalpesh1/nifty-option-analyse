param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot "nifty_option_data_cache.xlsx"),
  [int]$IntervalSeconds = 60,
  [switch]$NoUpload
)

$ErrorActionPreference = "Stop"

function Invoke-DatUploaderOnce {
  param([switch]$LocalOnly)

  $uploader = Join-Path $PSScriptRoot "run_excel_dat_upload_1min.ps1"
  if (!(Test-Path -LiteralPath $uploader)) {
    throw "Missing uploader script: $uploader"
  }

  $args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $uploader,
    "-ExcelPath", $ExcelPath,
    "-Once"
  )
  if ($LocalOnly) { $args += "-NoUpload" }

  & powershell @args
}

if (!(Test-Path -LiteralPath $ExcelPath)) {
  throw "Excel file not found: $ExcelPath"
}

$fullExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path
$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($fullExcelPath)

  Write-Host "Excel auto refresh/save started. Press Ctrl+C to stop."
  Write-Host "Workbook: $fullExcelPath"
  Write-Host "Every: $IntervalSeconds seconds"
  if ($NoUpload) {
    Write-Host "Upload: OFF, local DAT export only"
  } else {
    Write-Host "Upload: ON, DAT will publish to GitHub Pages"
  }
  Write-Host ""
  Write-Host "Keep Excel and this black window open."
  Write-Host "If you manually edit data, save is automatic every cycle."
  Write-Host ""

  while ($true) {
    try {
      Write-Host ("Refreshing and saving Excel at {0}" -f (Get-Date -Format "HH:mm:ss"))

      $workbook.RefreshAll()
      try { $excel.CalculateUntilAsyncQueriesDone() } catch {}
      try { $excel.CalculateFull() } catch {}
      $workbook.Save()

      Invoke-DatUploaderOnce -LocalOnly:$NoUpload
      Write-Host ("Cycle complete at {0}" -f (Get-Date -Format "HH:mm:ss"))
    } catch {
      Write-Warning $_.Exception.Message
    }

    Start-Sleep -Seconds $IntervalSeconds
  }
} finally {
  try {
    if ($workbook) { $workbook.Save(); $workbook.Close($true) | Out-Null }
  } catch {}
  try {
    if ($excel) { $excel.Quit() | Out-Null }
  } catch {}
  foreach ($obj in @($workbook, $excel)) {
    if ($obj) {
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
    }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
