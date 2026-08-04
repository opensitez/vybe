# vybe-test: powershell/event_actions/conditional_action
Register-EngineEvent -SourceIdentifier CondAction -Action { if ($Event.MessageData -eq 'go') { $Global.Cond = 'ok' } }
New-Event -SourceIdentifier CondAction -MessageData 'go'
if ($Global.Cond -ne 'ok') {
    Write-Host "FAIL: expected conditional action to set ok"
    exit 1
}
Unregister-Event -SourceIdentifier CondAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
