# vybe-test: powershell/event_subscriptions/subscription_action
Register-EngineEvent -SourceIdentifier SubAction -Action { $Global.ActionDone = 1 }
New-Event -SourceIdentifier SubAction
if ($Global.ActionDone -ne 1) {
    Write-Host "FAIL: expected action executed"
    exit 1
}
Unregister-Event -SourceIdentifier SubAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
