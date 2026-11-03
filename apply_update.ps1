param(
    [Parameter(Mandatory=$true)][string]$Target,
    [Parameter(Mandatory=$true)][string]$Staged,
    [Parameter(Mandatory=$true)][int]$WaitPid
)
$ErrorActionPreference = 'Stop'
try {
    Wait-Process -Id $WaitPid -Timeout 180 -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $Staged)) { throw 'Загруженный файл отсутствует.' }
    if ([IO.Path]::GetFullPath($Staged) -ne ([IO.Path]::GetFullPath($Target) + '.new')) {
        throw 'Неверный путь обновления.'
    }
    for ($try = 0; $try -lt 30; $try++) {
        try {
            Copy-Item -LiteralPath $Staged -Destination $Target -Force
            Remove-Item -LiteralPath $Staged -Force
            Start-Process -FilePath $Target -WorkingDirectory (Split-Path -Parent $Target)
            exit 0
        }
        catch {
            if ($try -eq 29) { throw }
            Start-Sleep -Seconds 1
        }
    }
}
catch {
    $message = $_.Exception.Message
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.MessageBox]::Show("Обновление Глазок не установлено: $message")
    exit 1
}
