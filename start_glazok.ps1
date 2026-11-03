$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$serverExe = Join-Path $PSScriptRoot 'Глазок Сервер.exe'
$clientExe = Join-Path $PSScriptRoot 'Глазок.exe'
$healthUrl = 'http://127.0.0.1:8765/health'
$useLocalServer = $true
$connectionFile = Join-Path $env:APPDATA 'Glasok\client_connection.json'
if (Test-Path -LiteralPath $connectionFile) {
    try {
        $saved = Get-Content -LiteralPath $connectionFile -Raw | ConvertFrom-Json
        if ($saved.address -and $saved.address -ne 'http://127.0.0.1:8765') {
            $useLocalServer = $false
        }
    }
    catch { Write-Host 'Cannot read server address; using local server.' }
}

if (-not (Test-Path -LiteralPath $clientExe)) {
    Write-Host "Missing file: $clientExe" -ForegroundColor Red
    exit 2
}

function Test-GlasokServer {
    try {
        $response = Invoke-RestMethod -Uri $healthUrl -Method Get -TimeoutSec 2
        return $response.ok -eq $true
    }
    catch {
        return $false
    }
}

if ($useLocalServer -and (Test-Path -LiteralPath $serverExe) -and -not (Test-GlasokServer)) {
    Write-Host 'Starting Glasok Server...'
    Start-Process -FilePath $serverExe -WorkingDirectory $PSScriptRoot

    $ready = $false
    for ($attempt = 1; $attempt -le 15; $attempt++) {
        Start-Sleep -Seconds 1
        if (Test-GlasokServer) {
            $ready = $true
            break
        }
    }

    if (-not $ready) { Write-Host 'Server is not ready. Open the server window and check its log.' }
}

Write-Host 'Starting Glasok...'
Start-Process -FilePath $clientExe -WorkingDirectory $PSScriptRoot
exit 0
