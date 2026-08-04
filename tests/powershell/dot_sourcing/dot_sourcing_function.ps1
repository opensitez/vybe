# vybe-test: powershell/dot_sourcing/dot_sourcing_function
$script = "$PWD/dot_sourcing_fn.ps1"
Set-Content -Path $script -Value 'function Get-Val { "OK" }'
. $script
if ((Get-Val) -ne 'OK') {
    Write-Host 'FAIL'
    exit 1
}
Remove-Item $script
Write-Host 'PASS'
exit 0
