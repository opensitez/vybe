# vybe-test: powershell/eventing/receive_event
New-Event -SourceIdentifier ReceiveEvent -MessageData 'ok'
$ev = Receive-Event -SourceIdentifier ReceiveEvent
if ($ev.SourceIdentifier -ne 'ReceiveEvent') {
    Write-Host "FAIL: expected ReceiveEvent"
    exit 1
}
Write-Host "PASS"
exit 0
