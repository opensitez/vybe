# vybe-test: powershell/dot_sourcing/dot_sourcing_variable
$script = "$PWD/dot_sourcing_var.ps1"
Set-Content -Path $script -Value '$global:DotSourced = "OK"'
. $script
if ($global:DotSourced -ne 'OK') {
    Write-Host 'FAIL'
    exit 1
}
Remove-Item $script
Write-Host 'PASS'
exit 0
