# vybe-test: powershell/warning_streams/warning_preference_continue
$WarningPreference = 'Continue'
Write-Warning 'warn'
if ($WarningPreference -ne 'Continue') {
    Write-Host "FAIL: expected warning continue"
    exit 1
}
Write-Host 'PASS'
exit 0
