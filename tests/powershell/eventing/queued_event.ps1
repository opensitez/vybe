# vybe-test: powershell/eventing/queued_event
New-Event -SourceIdentifier QueueEvent -MessageData 1
$event = Get-Event -SourceIdentifier QueueEvent
if ($event.MessageData -ne 1) {
    Write-Host "FAIL: expected message data 1"
    exit 1
}
Remove-Event -SourceIdentifier QueueEvent
Write-Host "PASS"
exit 0
