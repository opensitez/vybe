# vybe-test: powershell/debug_streams/write_debug_simple
$DebugPreference = 'Continue'
Write-Debug 'dbg'
Write-Host 'PASS'
exit 0
