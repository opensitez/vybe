# vybe-test: powershell/object_instantiation/new_list
$list = New-Object System.Collections.ArrayList
if ($list.Add(1) -ne 0) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
