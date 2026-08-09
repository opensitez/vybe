# vybe-test: powershell/constant_variables/constant_variable_in_function_scope
function Define-FuncConst {
    New-Variable -Name "FUNC_CONST" -Value 888 -Option Constant
    return $FUNC_CONST
}
$val = Define-FuncConst
if ($val -ne 888) {
    Write-Host "FAIL: function internal constant expected 888, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
