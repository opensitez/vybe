# vybe-test: powershell/indexing/array_index_by_variable
$arr = 5,6,7
$idx = 2
if ($arr[$idx] -ne 7) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
