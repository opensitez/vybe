# vybe-test: powershell/eventing/new_event
New-Event -SourceIdentifier TestEvent -MessageData 'x'
$events = Get-Event -SourceIdentifier TestEvent
if (-not $events) {
    Write-Host "FAIL: expected event in queue"
    exit 1
}
Remove-Event -SourceIdentifier TestEvent
Write-Host "PASS"
exit 0
