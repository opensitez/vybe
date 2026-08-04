# vybe-test: powershell/error_streams/write_error_object
Write-Error 'error'
if ($Error[0].Exception.Message -notlike '*error*') {
    Write-Host "FAIL: expected error message"
    exit 1
}
Write-Host 'PASS'
exit 0
