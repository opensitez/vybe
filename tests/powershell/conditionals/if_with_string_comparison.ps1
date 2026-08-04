# vybe-test: powershell/conditionals/if_with_string_comparison
if ('a' -eq 'a') {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
