# vybe-test: powershell/conditionals/if_with_subexpression
if ($(1 + 1) -eq 2) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
