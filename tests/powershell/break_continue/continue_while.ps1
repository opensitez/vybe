# vybe-test: powershell/break_continue/continue_while
$result = 0
$i = 0
while ($i -lt 3) { $i++; if ($i -eq 2) { continue }; $result += $i }
if ($result -eq 4) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
