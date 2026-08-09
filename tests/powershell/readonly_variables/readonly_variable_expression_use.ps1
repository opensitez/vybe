# vybe-test: powershell/readonly_variables/readonly_variable_expression_use
New-Variable -Name "RO_MULT" -Value 3 -Option ReadOnly
$res = 10 * $RO_MULT
if ($res -ne 30) {
    Write-Host "FAIL: ReadOnly variable expression multiplication expected 30, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
