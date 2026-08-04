# vybe-test: powershell/parentheses/grouping
$result = (1 + 2) * 3
if ($result -ne 9) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
