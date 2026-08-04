# vybe-test: powershell/event_actions/action_in_multiple_sources
Register-EngineEvent -SourceIdentifier MultiSourceAction -Action { $Global.Count += 1 }
New-Event -SourceIdentifier MultiSourceAction
if ($Global.Count -ne 1) {
    Write-Host "FAIL: expected action in multi-source"
    exit 1
}
Unregister-Event -SourceIdentifier MultiSourceAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
