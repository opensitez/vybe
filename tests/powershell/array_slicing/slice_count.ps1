# vybe-test: powershell/array_slicing/slice_count
$arr = 10,20,30,40
if (($arr[1..3].Count) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
