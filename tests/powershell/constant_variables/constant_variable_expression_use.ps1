# vybe-test: powershell/constant_variables/constant_variable_expression_use
New-Variable -Name "FACTOR" -Value 5 -Option Constant
$res = 1..3 | ForEach-Object { $_ * $FACTOR }
if ($res[0] -ne 5 -or $res[2] -ne 15) {
    Write-Host "FAIL: Constant variable in expression expected 5, 10, 15"
    exit 1
}
Write-Host "PASS"
exit 0
