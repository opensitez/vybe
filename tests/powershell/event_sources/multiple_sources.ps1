# vybe-test: powershell/event_sources/multiple_sources
New-Event -SourceIdentifier SourceA
New-Event -SourceIdentifier SourceB
$events = Get-Event
if ($events.Count -lt 2) {
    Write-Host "FAIL: expected at least two events"
    exit 1
}
Remove-Event -SourceIdentifier SourceA,SourceB
Write-Host "PASS"
exit 0
