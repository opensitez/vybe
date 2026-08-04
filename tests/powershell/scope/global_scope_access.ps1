# vybe-test: powershell/scope/global_scope_access
$global:myVar = 100
function Test-GlobalAccess {
    return $global:myVar
}
$result = Test-GlobalAccess
if ($result -ne 100) {
    Write-Host "FAIL: expected 100, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
