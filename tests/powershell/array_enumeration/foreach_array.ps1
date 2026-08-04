# vybe-test: powershell/array_enumeration/foreach_array
$arr = 1,2,3
$result = ''
foreach ($x in $arr) { $result += $x }
if ($result -eq '123') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
