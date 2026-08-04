# vybe-test: powershell/event_actions/action_with_variables
$x = 2
Register-EngineEvent -SourceIdentifier VarAction -Action { $Global.Value = $x * 2 }
New-Event -SourceIdentifier VarAction
if ($Global.Value -ne 4) {
    Write-Host "FAIL: expected action with variable to yield 4"
    exit 1
}
Unregister-Event -SourceIdentifier VarAction -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
