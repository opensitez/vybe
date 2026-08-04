# vybe-test: powershell/hashtable_methods/add
$h = @{}
$h.Add('a',1)
if ($h['a'] -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
