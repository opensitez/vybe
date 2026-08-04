# vybe-test: powershell/warning_streams/warning_action_invoke
Write-Warning 'warn'
if ($true -ne $true) {
    Write-Host "FAIL: dummy fail"
    exit 1
}
Write-Host 'PASS'
exit 0
