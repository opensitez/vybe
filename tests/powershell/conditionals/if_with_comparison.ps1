# vybe-test: powershell/conditionals/if_with_comparison
if (3 -gt 2) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
