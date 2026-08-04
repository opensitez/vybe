# vybe-test: powershell/error_streams/error_from_command
$ErrorActionPreference = 'Continue'
Get-Item nonexistent -ErrorAction SilentlyContinue
if ($Error.Count -lt 1) {
    Write-Host "FAIL: expected command error"
    exit 1
}
Write-Host 'PASS'
exit 0
