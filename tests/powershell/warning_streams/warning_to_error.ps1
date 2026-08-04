# vybe-test: powershell/warning_streams/warning_to_error
$WarningPreference = 'Stop'
try { Write-Warning 'warn' } catch { $caught = $true }
if (-not $caught) {
    Write-Host "FAIL: expected warning to stop"
    exit 1
}
Write-Host 'PASS'
exit 0
