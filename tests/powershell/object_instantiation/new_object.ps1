# vybe-test: powershell/object_instantiation/new_object
$obj = New-Object System.Text.StringBuilder
if ($obj.ToString() -ne '') { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
