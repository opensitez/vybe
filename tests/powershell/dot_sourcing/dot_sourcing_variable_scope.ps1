# vybe-test: powershell/dot_sourcing/dot_sourcing_variable_scope
$script = "$PWD/dot_sourcing_scope.ps1"
Set-Content -Path $script -Value '$scriptVar = "OK"'
. $script
if ($scriptVar -ne 'OK') {
    Write-Host 'FAIL'
    exit 1
}
Remove-Item $script
Write-Host 'PASS'
exit 0
