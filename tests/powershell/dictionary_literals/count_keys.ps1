# vybe-test: powershell/dictionary_literals/count_keys
$h = @{ a = 1; b = 2 }
if ($h.Count -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
