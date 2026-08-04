# vybe-test: powershell/aliasing/get_alias
Set-Alias hi Write-Output
$alias = Get-Alias hi
if ($alias.Name -ne 'hi') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
