# vybe-test: powershell/expressions/subexpression_expression
if ($(1 + 1) -eq 2) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
