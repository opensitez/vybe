# vybe-test: powershell/logical_operators/or_with_comparison
if ((1 -eq 2) -or (2 -eq 2)) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
