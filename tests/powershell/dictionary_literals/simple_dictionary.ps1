# vybe-test: powershell/dictionary_literals/simple_dictionary
$h = @{ a = 1; b = 2 }
if ($h.a -eq 1 -and $h.b -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
