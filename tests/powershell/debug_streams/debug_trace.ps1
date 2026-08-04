# vybe-test: powershell/debug_streams/debug_trace
$DebugPreference = 'Continue'
Write-Debug 'trace'
Write-Host 'PASS'
exit 0
