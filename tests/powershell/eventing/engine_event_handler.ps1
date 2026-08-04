# vybe-test: powershell/eventing/engine_event_handler
Register-EngineEvent -SourceIdentifier MyEngineEvent -Action { $Global.Handled = 1 }
New-Event -SourceIdentifier MyEngineEvent
if ($Global.Handled -ne 1) {
    Write-Host "FAIL: expected handler run"
    exit 1
}
Unregister-Event -SourceIdentifier MyEngineEvent -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
