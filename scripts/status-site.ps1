$ErrorActionPreference = "Stop"

$Port = 5500
$connections = Get-NetTCPConnection -LocalAddress 127.0.0.1 -LocalPort $Port -State Listen -ErrorAction SilentlyContinue

if (-not $connections) {
  Write-Host "Website server is stopped."
  exit 0
}

foreach ($connection in $connections) {
  $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
  Write-Host "Website server is running: http://127.0.0.1:$Port/index.html"
  if ($process) {
    Write-Host "PID: $($process.Id) ($($process.ProcessName))"
  }
}
