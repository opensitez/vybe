# vybe-test: powershell/logical_operators/not_with_comparison
if (-not (1 -eq 2)) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
