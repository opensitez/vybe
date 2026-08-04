# vybe-test: powershell/event_subscriptions/subscription_source_name
Register-EngineEvent -SourceIdentifier NameSub -Action { $Global.NameSub = $true }
New-Event -SourceIdentifier NameSub
if (-not $Global.NameSub) {
    Write-Host "FAIL: expected event source name match"
    exit 1
}
Unregister-Event -SourceIdentifier NameSub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
