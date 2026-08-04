# vybe-test: powershell/object_instantiation/new_hashtable
$hash = @{}
$hash.a = 1
if ($hash.a -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
