# vybe-test: powershell/type_checking/is_hashtable
if (@{a=1} -is [hashtable]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
