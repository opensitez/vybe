# vybe-test: powershell/hashtable_methods/foreach
$h = @{ a=1 }
foreach ($k in $h.Keys) { if ($k -eq 'a') { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
