# vybe-test: powershell/hashtable_methods/clear
$h = @{ a=1 }
$h.Clear()
if ($h.Count -eq 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
