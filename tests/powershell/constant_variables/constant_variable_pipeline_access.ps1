# vybe-test: powershell/constant_variables/constant_variable_pipeline_access
New-Variable -Name "PIPE_CONST" -Value @(10, 20) -Option Constant
$sum = ($PIPE_CONST | Measure-Object -Sum).Sum
if ($sum -ne 30) {
    Write-Host "FAIL: Constant variable pipeline Measure-Object expected 30, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
