# vybe-test: powershell/dictionary_literals/add_key
$h = @{ }
$h['a'] = 1
if ($h.a -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
