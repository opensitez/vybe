# vybe-test: powershell/error_streams/error_action_continue
$ErrorActionPreference = 'Continue'
Write-Error 'err'
if ($Error.Count -lt 1) {
    Write-Host "FAIL: expected error logged"
    exit 1
}
Write-Host 'PASS'
exit 0
