# vybe-test: powershell/debug_streams/debug_clear
$DebugPreference = 'Continue'
$DebugPreference = 'SilentlyContinue'
if ($DebugPreference -ne 'SilentlyContinue') {
    Write-Host "FAIL: expected clear with preference"
    exit 1
}
Write-Host 'PASS'
exit 0
