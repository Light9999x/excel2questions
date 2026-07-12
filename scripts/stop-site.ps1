$ErrorActionPreference = "Stop"

$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$Port = 5500
$PidFile = Join-Path $ProjectRoot ".site-server.pid"
$stopped = $false

if (Test-Path -LiteralPath $PidFile) {
  $pidText = Get-Content -LiteralPath $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($pidText -match "^\d+$") {
    $process = Get-Process -Id ([int]$pidText) -ErrorAction SilentlyContinue
    if ($process) {
      Stop-Process -Id $process.Id
      Write-Host "Stopped website server PID: $($process.Id)"
      $stopped = $true
    }
  }
  Remove-Item -LiteralPath $PidFile -Force
}

$connections = Get-NetTCPConnection -LocalAddress 127.0.0.1 -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
foreach ($connection in $connections) {
  $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
  if ($process -and $process.ProcessName -match "python") {
    Stop-Process -Id $process.Id
    Write-Host "Stopped website server PID: $($process.Id)"
    $stopped = $true
  }
}

if (-not $stopped) {
  Write-Host "Website server is not running."
}
