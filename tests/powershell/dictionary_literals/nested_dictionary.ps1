# vybe-test: powershell/dictionary_literals/nested_dictionary
$h = @{ a = @{ b = 2 } }
if ($h.a.b -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
