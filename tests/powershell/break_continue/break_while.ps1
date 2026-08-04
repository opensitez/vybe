# vybe-test: powershell/break_continue/break_while
$result = 0
$i = 0
while ($i -lt 5) { if ($i -eq 2) { break }; $result += $i; $i++ }
if ($result -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
