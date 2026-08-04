# vybe-test: powershell/indexing/subexpression_indexing
$arr = 1,2,3
$index = 2
if ($arr[$index] -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
