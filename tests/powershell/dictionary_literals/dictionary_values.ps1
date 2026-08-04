# vybe-test: powershell/dictionary_literals/dictionary_values
$h = @{ a=1 }
if ($h.Values -contains 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
