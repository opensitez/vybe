# vybe-test: powershell/hashtable_methods/keys
$h = @{ a=1; b=2 }
if (($h.Keys -join ',') -match 'a') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
