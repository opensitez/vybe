# vybe-test: powershell/hashtable_methods/containsvalue
$h = @{ a=1 }
if ($h.ContainsValue(1)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
