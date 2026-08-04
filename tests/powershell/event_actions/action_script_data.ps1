# vybe-test: powershell/event_actions/action_script_data
Register-EngineEvent -SourceIdentifier ScriptData -Action { param($e) $Global.Data = $e.MessageData }
New-Event -SourceIdentifier ScriptData -MessageData @{ A = 9 }
if ($Global.Data.A -ne 9) {
    Write-Host "FAIL: expected action message data 9"
    exit 1
}
Unregister-Event -SourceIdentifier ScriptData -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
