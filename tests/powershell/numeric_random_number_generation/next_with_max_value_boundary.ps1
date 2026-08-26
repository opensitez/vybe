# vybe-test: powershell/numeric_random_number_generation/next_with_max_value_boundary
$rnd = [System.Random]::new()
for ($i = 0; $i -lt 50; $i++) {
    $v = $rnd.Next(10)
    if ($v -lt 0 -or $v -ge 10) {
        Write-Host "FAIL: Next(10) out of [0, 10) range: $v"
        exit 1
    }
}
Write-Host "PASS"
exit 0
