# vybe-test: powershell/array_enumeration/while_array
$arr = 1,2,3
$i = 0
$result = 0
while ($i -lt $arr.Count) { $result += $arr[$i]; $i++ }
if ($result -eq 6) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
