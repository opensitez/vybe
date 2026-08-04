# vybe-test: powershell/error_streams/error_action_stop
$ErrorActionPreference = 'Stop'
try {
    Write-Error 'err' -ErrorAction Stop
    $caught = $false
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: expected error catch"
    exit 1
}
Write-Host 'PASS'
exit 0
