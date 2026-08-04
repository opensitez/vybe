# vybe-test: powershell/output_streams/output_stream_string
$result = Write-Output 'string'
if ($result -ne 'string') {
    Write-Host "FAIL: expected string"
    exit 1
}
Write-Host 'PASS'
exit 0
