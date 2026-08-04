# vybe-test: powershell/array_enumeration/array_element_sum
$arr = 1,2,3
if (($arr | Measure-Object -Sum).Sum -eq 6) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
