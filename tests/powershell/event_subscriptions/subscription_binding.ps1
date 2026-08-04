# vybe-test: powershell/event_subscriptions/subscription_binding
Register-EngineEvent -SourceIdentifier BoundSub -Action { $Global.Bound = $true }
New-Event -SourceIdentifier BoundSub
if (-not $Global.Bound) {
    Write-Host "FAIL: expected bound handler"
    exit 1
}
Unregister-Event -SourceIdentifier BoundSub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
