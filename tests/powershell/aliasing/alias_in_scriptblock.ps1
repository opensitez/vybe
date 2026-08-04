# vybe-test: powershell/aliasing/alias_in_scriptblock
Set-Alias hi Write-Output
& { hi 'PASS' }
exit 0
