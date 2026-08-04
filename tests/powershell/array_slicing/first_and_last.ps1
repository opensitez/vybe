# vybe-test: powershell/array_slicing/first_and_last
$arr = 5,6,7
if ($arr[0] -eq 5 -and $arr[2] -eq 7) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
