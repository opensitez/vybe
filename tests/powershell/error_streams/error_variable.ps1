# vybe-test: powershell/error_streams/error_variable
Write-Error 'err'
if (-not $Error) {
    Write-Host "FAIL: expected $Error populated"
    exit 1
}
Write-Host 'PASS'
exit 0
