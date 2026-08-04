# vybe-test: powershell/writer_streams/output_writer
Write-Output 'x'
if ($true -ne $true) {
    Write-Host "FAIL: dummy fail"
    exit 1
}
Write-Host 'PASS'
exit 0
