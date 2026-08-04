# vybe-test: powershell/event_subscriptions/engine_event_subscription
Register-EngineEvent -SourceIdentifier SubEvent -Action { $Global.Handled = 'yes' }
New-Event -SourceIdentifier SubEvent
if ($Global.Handled -ne 'yes') {
    Write-Host "FAIL: expected subscription handled"
    exit 1
}
Unregister-Event -SourceIdentifier SubEvent -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
