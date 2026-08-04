# vybe-test: powershell/event_actions/action_scriptblock
Register-EngineEvent -SourceIdentifier ActionEvent -Action { $Global.ActionDone = 5 }
New-Event -SourceIdentifier ActionEvent
if ($Global.ActionDone -ne 5) {
    Write-Host "FAIL: expected action scriptblock run"
    exit 1
}
Unregister-Event -SourceIdentifier ActionEvent -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
