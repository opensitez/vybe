# vybe-test: powershell/break_continue/continue_for
$result = 0
for ($i=0; $i -lt 3; $i++) { if ($i -eq 1) { continue }; $result += $i }
if ($result -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
