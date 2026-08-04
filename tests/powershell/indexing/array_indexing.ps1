# vybe-test: powershell/indexing/array_indexing
$arr = 1,2,3
if ($arr[1] -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
