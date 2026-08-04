# vybe-test: powershell/break_continue/break_nested
$result = 0
foreach ($x in 1,2) { foreach ($y in 1,2) { if ($y -eq 2) { break }; $result += $y } }
if ($result -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
