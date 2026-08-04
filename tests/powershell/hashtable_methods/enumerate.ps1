# vybe-test: powershell/hashtable_methods/enumerate
$h = @{ a=1; b=2 }
if (($h.GetEnumerator() | Where-Object { $_.Key -eq 'a' }).Value -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
