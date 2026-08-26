# vybe-test: powershell/numeric_random_number_generation/get_random_powershell_cmdlet_basic
$num = Get-Random -Minimum 1 -Maximum 10
if ($num -lt 1 -or $num -ge 10) {
    Write-Host "FAIL: Get-Random cmdlet out of range: $num"
    exit 1
}
Write-Host "PASS"
exit 0
