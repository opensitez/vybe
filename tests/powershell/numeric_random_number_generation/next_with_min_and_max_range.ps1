# vybe-test: powershell/numeric_random_number_generation/next_with_min_and_max_range
$rnd = [System.Random]::new()
for ($i = 0; $i -lt 50; $i++) {
    $v = $rnd.Next(50, 60)
    if ($v -lt 50 -or $v -ge 60) {
        Write-Host "FAIL: Next(50, 60) out of [50, 60) range: $v"
        exit 1
    }
}
Write-Host "PASS"
exit 0
