# vybe-test: powershell/parentheses/array_expression
$arr = @(1, (2 + 3), 4)
if ($arr[1] -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
