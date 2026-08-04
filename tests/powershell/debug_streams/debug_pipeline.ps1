# vybe-test: powershell/debug_streams/debug_pipeline
$DebugPreference = 'Continue'
1..2 | ForEach-Object { Write-Debug 'd' }
Write-Host 'PASS'
exit 0
