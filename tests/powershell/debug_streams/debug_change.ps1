# vybe-test: powershell/debug_streams/debug_change
$DebugPreference = 'Continue'
if ($DebugPreference -ne 'Continue') {
    Write-Host "FAIL: expected changed preference"
    exit 1
}
Write-Host 'PASS'
exit 0
