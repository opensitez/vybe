# vybe-test: powershell/numeric_random_number_generation/get_random_with_seed_is_deterministic
$v1 = Get-Random -Minimum 1 -Maximum 1000 -SetSeed 54321
$v2 = Get-Random -Minimum 1 -Maximum 1000 -SetSeed 54321
if ($v1 -ne $v2) {
    Write-Host "FAIL: Get-Random with -SetSeed failed"
    exit 1
}
Write-Host "PASS"
exit 0
