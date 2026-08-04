# vybe-test: powershell/scope/script_scope
$script:value = 42
function Get-ScriptValue {
    return $script:value
}
$result = Get-ScriptValue
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
