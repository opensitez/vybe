# vybe-test: powershell/parentheses/nested_parentheses
if ((1 + (2 * 3)) -eq 7) {
    Write-Output 'PASS'
}
exit 0
