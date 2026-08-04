# vybe-test: powershell/break_continue/break_do
$result = 0
do { if ($result -eq 1) { break }; $result++ } while ($result -lt 5)
if ($result -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
