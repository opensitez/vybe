# vybe-test: powershell/object_instantiation/new_type
$type = [int]
if ($type -ne [int]) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
