# vybe-test: powershell/array_initialization/array_from_expression
$arr = @(1 + 1, 2 + 2)
if ($arr[0] -eq 2 -and $arr[1] -eq 4) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
