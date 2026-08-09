# vybe-test: powershell/constant_variables/constant_variable_string_type
New-Variable -Name "STR_CONST" -Value "ConstString" -Option Constant
if ($STR_CONST.Length -ne 11) {
    Write-Host "FAIL: Constant string length expected 11, got $($STR_CONST.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
