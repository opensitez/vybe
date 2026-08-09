# vybe-test: powershell/readonly_variables/readonly_variable_in_function_scope
function Set-ROLocal {
    New-Variable -Name "LOCAL_RO" -Value "FnScope" -Option ReadOnly
    return $LOCAL_RO
}
$res = Set-ROLocal
if ($res -ne "FnScope") {
    Write-Host "FAIL: function scope ReadOnly variable expected FnScope, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
