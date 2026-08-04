# vybe-test: powershell/dictionary_literals/remove_key
$h = @{ a = 1 }
$h.Remove('a') | Out-Null
if (-not $h.ContainsKey('a')) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
