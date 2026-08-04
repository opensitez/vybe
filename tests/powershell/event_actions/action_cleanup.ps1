# vybe-test: powershell/event_actions/action_cleanup
Register-EngineEvent -SourceIdentifier CleanupAction -Action { }
Unregister-Event -SourceIdentifier CleanupAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
