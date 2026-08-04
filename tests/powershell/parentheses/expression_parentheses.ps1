# vybe-test: powershell/parentheses/expression_parentheses
if ((1 + 2) * (3 - 1) -eq 6) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
