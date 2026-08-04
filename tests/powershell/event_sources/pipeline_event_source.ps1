# vybe-test: powershell/event_sources/pipeline_event_source
1..3 | ForEach-Object { New-Event -SourceIdentifier PipeSource -MessageData $_ }
$events = Get-Event -SourceIdentifier PipeSource
if ($events.Count -ne 3) {
    Write-Host "FAIL: expected three events"
    exit 1
}
Remove-Event -SourceIdentifier PipeSource
Write-Host "PASS"
exit 0
