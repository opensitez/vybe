# vybe-test: powershell/array_slicing/array_slice_in_expression
$arr = 1,2,3,4
if (($arr[0..1] | Measure-Object).Count -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
