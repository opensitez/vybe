# vybe-test: powershell/object_instantiation/new_guid
$guid = [guid]::NewGuid()
if ($guid -eq $null) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
