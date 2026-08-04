# vybe-test: powershell/object_instantiation/new_regex
$regex = [regex]'a'
if ($regex -eq $null) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
