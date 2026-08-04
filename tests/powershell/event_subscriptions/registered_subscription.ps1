# vybe-test: powershell/event_subscriptions/registered_subscription
$subscription = Register-EngineEvent -SourceIdentifier SubRegistered -Action { $Global.Flag = 1 }
if (-not $subscription) {
    Write-Host "FAIL: expected subscription object"
    exit 1
}
Unregister-Event -SourceIdentifier SubRegistered -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
