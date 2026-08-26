# vybe-test: powershell/numeric_random_number_generation/get_random_double_range
$val = Get-Random -Minimum 0.5 -Maximum 2.5
if ($val -lt 0.5 -or $val -ge 2.5) {
    Write-Host "FAIL: Get-Random double range failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
