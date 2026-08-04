# vybe-test: powershell/null_handling/null_in_hashtable
$h = @{ a = $null }
if ($h.a -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
