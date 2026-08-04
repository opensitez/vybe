# vybe-test: powershell/array_slicing/expression_index
$arr = 1,2,3,4
if ($arr[(1+1)] -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
