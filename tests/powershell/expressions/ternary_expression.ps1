# vybe-test: powershell/expressions/ternary_expression
if ((1 -eq 1) ? $true : $false) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
