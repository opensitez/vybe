# vybe-test: powershell/writer_streams/error_writer
Write-Error 'test'
if ($Error[0].ToString() -notlike '*test*') {
    Write-Host "FAIL: expected error object"
    exit 1
}
Write-Host 'PASS'
exit 0
