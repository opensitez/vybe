# vybe-test: powershell/debug_streams/debug_multiple
$DebugPreference = 'Continue'
Write-Debug 'a'
Write-Debug 'b'
Write-Host 'PASS'
exit 0
