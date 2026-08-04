# vybe-test: powershell/dictionary_literals/dictionary_to_array
$h = @{ a=1; b=2 }
if (($h.GetEnumerator() | Measure-Object).Count -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
