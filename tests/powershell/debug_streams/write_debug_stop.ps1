# vybe-test: powershell/debug_streams/write_debug_stop
$DebugPreference = 'Stop'
try { Write-Debug 'dbg' } catch { $caught = $true }
Write-Host 'PASS'
exit 0
