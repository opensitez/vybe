# vybe-test: powershell/warning_streams/warning_default
if ($WarningPreference -ne 'Continue') {
    Write-Host "FAIL: expected default WarningPreference"
    exit 1
}
Write-Host 'PASS'
exit 0
