# vybe-test: powershell/hashtable_methods/values
$h = @{ a=1; b=2 }
if (($h.Values -join ',') -match '2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
