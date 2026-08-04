# vybe-test: powershell/subexpressions/subexpression_comparison
if ($(1 + 1) -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
