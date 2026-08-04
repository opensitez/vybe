# vybe-test: powershell/expressions/negated_expression
if (-not (1 -eq 2)) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
