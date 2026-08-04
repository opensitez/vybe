# vybe-test: powershell/dot_sourcing/dot_sourcing_array
$script = "$PWD/dot_sourcing_arr.ps1"
Set-Content -Path $script -Value '$arr = 1,2,3'
. $script
if ($arr[1] -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Remove-Item $script
Write-Host 'PASS'
exit 0
