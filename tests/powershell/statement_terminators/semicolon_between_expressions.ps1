# vybe-test: powershell/statement_terminators/semicolon_between_expressions
$x = 1; $y = 2
if ($x -eq 1 -and $y -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
