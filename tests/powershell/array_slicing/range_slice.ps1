# vybe-test: powershell/array_slicing/range_slice
$arr = 1,2,3,4
if (($arr[1..2] -join ',') -eq '2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
