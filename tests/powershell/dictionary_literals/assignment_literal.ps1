# vybe-test: powershell/dictionary_literals/assignment_literal
$h = @{ a = 1 }
if ($h.a -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
