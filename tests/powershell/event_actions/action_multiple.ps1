# vybe-test: powershell/event_actions/action_multiple
Register-EngineEvent -SourceIdentifier MultiAction -Action { $Global.Count += 1 }
Register-EngineEvent -SourceIdentifier MultiAction -Action { $Global.Count += 2 }
New-Event -SourceIdentifier MultiAction
if ($Global.Count -ne 3) {
    Write-Host "FAIL: expected combined action results"
    exit 1
}
Unregister-Event -SourceIdentifier MultiAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
