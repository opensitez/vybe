# vybe-test: powershell/event_actions/delayed_action
Register-EngineEvent -SourceIdentifier DelayedAction -Action { $Global.Delayed = 1 }
Start-Sleep -Milliseconds 10
New-Event -SourceIdentifier DelayedAction
if ($Global.Delayed -ne 1) {
    Write-Host "FAIL: expected delayed action to run"
    exit 1
}
Unregister-Event -SourceIdentifier DelayedAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
