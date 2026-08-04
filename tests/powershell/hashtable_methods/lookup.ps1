# vybe-test: powershell/hashtable_methods/lookup
$h = @{ a=1 }
if ($h['a'] -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
