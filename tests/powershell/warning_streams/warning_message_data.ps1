# vybe-test: powershell/warning_streams/warning_message_data
Write-Warning 'data'
if ($true -ne $true) {
    Write-Host "FAIL: dummy fail"
    exit 1
}
Write-Host 'PASS'
exit 0
