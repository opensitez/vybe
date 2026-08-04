# vybe-test: powershell/warning_streams/warning_preference_silentlycontinue
$WarningPreference = 'SilentlyContinue'
Write-Warning 'warn'
if ($WarningPreference -ne 'SilentlyContinue') {
    Write-Host "FAIL: expected silent warning"
    exit 1
}
Write-Host 'PASS'
exit 0
