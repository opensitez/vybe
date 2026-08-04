# vybe-test: powershell/event_actions/action_parameter
Register-EngineEvent -SourceIdentifier ParamAction -Action { param($e) $Global.Param = $e.MessageData }
New-Event -SourceIdentifier ParamAction -MessageData 7
if ($Global.Param -ne 7) {
    Write-Host "FAIL: expected action parameter 7"
    exit 1
}
Unregister-Event -SourceIdentifier ParamAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
