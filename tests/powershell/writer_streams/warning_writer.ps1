# vybe-test: powershell/writer_streams/warning_writer
Write-Warning 'warn'
if ($true -ne $true) {
    Write-Host "FAIL: dummy fail"
    exit 1
}
Write-Host 'PASS'
exit 0
