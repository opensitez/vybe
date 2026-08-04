# vybe-test: powershell/debug_streams/debug_note_property
$DebugPreference = 'Continue'
Write-Debug ([PSCustomObject]@{ A='x' })
Write-Host 'PASS'
exit 0
