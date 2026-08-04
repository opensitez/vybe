# vybe-test: powershell/array_enumeration/array_enumerate_nested
$arr = @(1,@(2,3))
$result = ''
foreach ($x in $arr) { if ($x -is [array]) { foreach ($y in $x) { $result += $y } } else { $result += $x } }
if ($result -eq '123') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
