# vybe-test: powershell/eventing/clear_event_queue
New-Event -SourceIdentifier ClearEvent
Remove-Event -SourceIdentifier ClearEvent
$ev = Get-Event -SourceIdentifier ClearEvent -ErrorAction SilentlyContinue
if ($ev) {
    Write-Host "FAIL: expected empty queue"
    exit 1
}
Write-Host "PASS"
exit 0
