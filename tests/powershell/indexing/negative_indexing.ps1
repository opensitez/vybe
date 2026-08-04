# vybe-test: powershell/indexing/negative_indexing
$arr = 1,2,3
if ($arr[-1] -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
