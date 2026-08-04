# vybe-test: powershell/warning_streams/warning_pref_change
$old = $WarningPreference
$WarningPreference = 'SilentlyContinue'
if ($WarningPreference -ne 'SilentlyContinue') {
    Write-Host "FAIL: expected preference changed"
    exit 1
}
$WarningPreference = $old
Write-Host 'PASS'
exit 0
