# vybe-test: powershell/expressions/combined_expression
if ((1 + 1) -eq 2 -and (2 -eq 2)) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
