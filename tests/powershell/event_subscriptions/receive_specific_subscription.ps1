# vybe-test: powershell/event_subscriptions/receive_specific_subscription
Register-EngineEvent -SourceIdentifier ReceiveSub -Action { }
New-Event -SourceIdentifier ReceiveSub
$event = Receive-Event -SourceIdentifier ReceiveSub
if ($event.SourceIdentifier -ne 'ReceiveSub') {
    Write-Host "FAIL: expected ReceiveSub event"
    exit 1
}
Unregister-Event -SourceIdentifier ReceiveSub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
