# vybe-test: powershell/expressions/comparison_expression
if (2 -gt 1) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
