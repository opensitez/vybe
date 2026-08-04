# vybe-test: powershell/hashtable_methods/containskey
$h = @{ a=1 }
if ($h.ContainsKey('a')) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
