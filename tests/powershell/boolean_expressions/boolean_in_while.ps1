# vybe-test: powershell/boolean_expressions/boolean_in_while
$i = 0
while ($i -lt 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
