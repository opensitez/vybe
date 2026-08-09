# vybe-test: powershell/constant_variables/constant_variable_array_type
New-Variable -Name "ARR_CONST" -Value @("X", "Y", "Z") -Option Constant
if ($ARR_CONST.Count -ne 3 -or $ARR_CONST[1] -ne "Y") {
    Write-Host "FAIL: Constant array expected Count 3, item 'Y'"
    exit 1
}
Write-Host "PASS"
exit 0
