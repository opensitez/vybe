# vybe-test: powershell/constant_variables/constant_variable_set
New-Variable -Name "MY_CONST" -Value 100 -Option Constant
if ($MY_CONST -ne 100) {
    Write-Host "FAIL: Constant variable creation expected 100, got $MY_CONST"
    exit 1
}
Write-Host "PASS"
exit 0
