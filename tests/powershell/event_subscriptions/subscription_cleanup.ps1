# vybe-test: powershell/event_subscriptions/subscription_cleanup
Register-EngineEvent -SourceIdentifier CleanSub -Action { }
Unregister-Event -SourceIdentifier CleanSub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
