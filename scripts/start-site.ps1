$ErrorActionPreference = "Stop"

$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$Port = 5500
$PidFile = Join-Path $ProjectRoot ".site-server.pid"

$existing = Get-NetTCPConnection -LocalAddress 127.0.0.1 -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($existing) {
  Write-Host "Website is already running: http://127.0.0.1:$Port/index.html"
  exit 0
}

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
  Write-Error "Python was not found. Install Python or change this script to use another static server."
  exit 1
}

$process = Start-Process `
  -FilePath $python.Source `
  -ArgumentList @("-m", "http.server", "$Port", "--bind", "127.0.0.1") `
  -WorkingDirectory $ProjectRoot `
  -WindowStyle Hidden `
  -PassThru

Set-Content -LiteralPath $PidFile -Value $process.Id
Write-Host "Website started: http://127.0.0.1:$Port/index.html"
Write-Host "PID: $($process.Id)"
