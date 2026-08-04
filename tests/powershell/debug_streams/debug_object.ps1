# vybe-test: powershell/debug_streams/debug_object
$DebugPreference = 'Continue'
Write-Debug ([PSCustomObject]@{ Value = 1 })
Write-Host 'PASS'
exit 0
