# vybe-test: powershell/constant_variables/constant_variable_int_type
New-Variable -Name "INT_CONST" -Value ([int]123) -Option Constant
if (-not ($INT_CONST -is [int])) {
    Write-Host "FAIL: Constant variable type expected int"
    exit 1
}
Write-Host "PASS"
exit 0
