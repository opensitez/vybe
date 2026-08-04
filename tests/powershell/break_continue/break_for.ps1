# vybe-test: powershell/break_continue/break_for
$result = 0
for ($i=0; $i -lt 5; $i++) { if ($i -eq 2) { break }; $result += $i }
if ($result -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
