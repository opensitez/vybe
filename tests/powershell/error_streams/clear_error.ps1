# vybe-test: powershell/error_streams/clear_error
Write-Error 'x'
$Error.Clear()
if ($Error.Count -ne 0) {
    Write-Host "FAIL: expected error cleared"
    exit 1
}
Write-Host 'PASS'
exit 0
