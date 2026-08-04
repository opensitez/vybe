# vybe-test: powershell/error_streams/error_count
$ErrorActionPreference = 'Continue'
Write-Error 'one'
Write-Error 'two'
if ($Error.Count -lt 2) {
    Write-Host "FAIL: expected two errors"
    exit 1
}
Write-Host 'PASS'
exit 0
