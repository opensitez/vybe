# vybe-test: powershell/dictionary_literals/lookup_value
$h = @{ a = 1 }
if ($h['a'] -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
