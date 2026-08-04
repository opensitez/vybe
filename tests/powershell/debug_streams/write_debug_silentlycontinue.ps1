# vybe-test: powershell/debug_streams/write_debug_silentlycontinue
$DebugPreference = 'SilentlyContinue'
Write-Debug 'dbg'
if ($DebugPreference -ne 'SilentlyContinue') {
    Write-Host "FAIL: expected SilentlyContinue"
    exit 1
}
Write-Host 'PASS'
exit 0
