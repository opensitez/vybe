# vybe-test: powershell/debug_streams/debug_preference
$DebugPreference = 'Continue'
if ($DebugPreference -ne 'Continue') {
    Write-Host "FAIL: expected Continue"
    exit 1
}
Write-Host 'PASS'
exit 0
