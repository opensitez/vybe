# vybe-test: powershell/type_operators/is_type_hashtable
if (@{a=1} -is [hashtable]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
