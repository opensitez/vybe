# vybe-test: powershell/break_continue/break_foreach
$result = 0
foreach ($x in 1,2,3) { if ($x -eq 2) { break }; $result += $x }
if ($result -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
