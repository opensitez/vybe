# vybe-test: powershell/break_continue/continue_nested
$result = 0
foreach ($x in 1,2) { foreach ($y in 1,2) { if ($y -eq 2) { continue }; $result += $y } }
if ($result -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
