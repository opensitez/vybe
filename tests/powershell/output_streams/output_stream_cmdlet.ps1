# vybe-test: powershell/output_streams/output_stream_cmdlet
$result = Write-Output 'x'
if ($result -ne 'x') {
    Write-Host "FAIL: expected x"
    exit 1
}
Write-Host 'PASS'
exit 0
