# vybe-test: powershell/readonly_variables/readonly_variable_scope_inheritance
New-Variable -Name "PARENT_RO" -Value "ScopeVal" -Option ReadOnly
function Test-ROScope {
    return $PARENT_RO
}
$res = Test-ROScope
if ($res -ne "ScopeVal") {
    Write-Host "FAIL: child scope ReadOnly variable reading expected ScopeVal, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
