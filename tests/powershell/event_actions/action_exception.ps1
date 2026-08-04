# vybe-test: powershell/event_actions/action_exception
Register-EngineEvent -SourceIdentifier ActionEx -Action { throw 'boom' }
New-Event -SourceIdentifier ActionEx
Unregister-Event -SourceIdentifier ActionEx -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
