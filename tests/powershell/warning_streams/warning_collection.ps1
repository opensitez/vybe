# vybe-test: powershell/warning_streams/warning_collection
$WarningPreference = 'Continue'
Write-Warning 'warn'
if ($WarningPreference -ne 'Continue') {
    Write-Host "FAIL: expected continue"
    exit 1
}
Write-Host 'PASS'
exit 0
