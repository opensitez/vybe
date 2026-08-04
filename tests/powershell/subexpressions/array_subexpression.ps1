# vybe-test: powershell/subexpressions/array_subexpression
$value = @(1,2,3)
if ($value.Count -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
