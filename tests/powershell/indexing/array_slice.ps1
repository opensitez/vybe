# vybe-test: powershell/indexing/array_slice
$arr = 1,2,3,4
if (($arr[1..2] -join ',') -ne '2,3') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
