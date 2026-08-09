# vybe-test: powershell/constant_variables/constant_variable_scope_inheritance
New-Variable -Name "PARENT_CONST" -Value "ParentVal" -Option Constant
function Child-ReadConst {
    return $PARENT_CONST
}
$res = Child-ReadConst
if ($res -ne "ParentVal") {
    Write-Host "FAIL: Constant variable child scope reading expected ParentVal, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
