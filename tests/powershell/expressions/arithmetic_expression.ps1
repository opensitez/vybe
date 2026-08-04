# vybe-test: powershell/expressions/arithmetic_expression
if ((1 + 2) -eq 3) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
