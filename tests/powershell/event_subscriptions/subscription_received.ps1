# vybe-test: powershell/event_subscriptions/subscription_received
Register-EngineEvent -SourceIdentifier ReceivedSub -Action { $Global.Received = $true }
New-Event -SourceIdentifier ReceivedSub
if (-not $Global.Received) {
    Write-Host "FAIL: expected event delivery"
    exit 1
}
Unregister-Event -SourceIdentifier ReceivedSub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
