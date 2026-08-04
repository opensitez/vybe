# vybe-test: powershell/break_continue/continue_foreach
$result = 0
foreach ($x in 1,2,3) { if ($x -eq 2) { continue }; $result += $x }
if ($result -eq 4) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
