# vybe-test: powershell/event_sources/event_source_type
$event = New-Event -SourceIdentifier TypeSource -MessageData 3
if ($event.SourceIdentifier -ne 'TypeSource') {
    Write-Host "FAIL: expected TypeSource"
    exit 1
}
Remove-Event -SourceIdentifier TypeSource
Write-Host "PASS"
exit 0
