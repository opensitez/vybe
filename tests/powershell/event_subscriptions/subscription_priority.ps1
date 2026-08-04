# vybe-test: powershell/event_subscriptions/subscription_priority
Register-EngineEvent -SourceIdentifier PrioritySub -Action { $Global.Priority += 1 }
New-Event -SourceIdentifier PrioritySub
if ($Global.Priority -ne 1) {
    Write-Host "FAIL: expected single invocation"
    exit 1
}
Unregister-Event -SourceIdentifier PrioritySub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
