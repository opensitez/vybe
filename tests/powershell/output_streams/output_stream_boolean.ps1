# vybe-test: powershell/output_streams/output_stream_boolean
$result = Write-Output $true
if ($result -ne $true) {
    Write-Host "FAIL: expected true"
    exit 1
}
Write-Host 'PASS'
exit 0
