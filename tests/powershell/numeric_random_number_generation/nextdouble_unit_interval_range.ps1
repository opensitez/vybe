# vybe-test: powershell/numeric_random_number_generation/nextdouble_unit_interval_range
$rnd = [System.Random]::new()
for ($i = 0; $i -lt 50; $i++) {
    $d = $rnd.NextDouble()
    if ($d -lt 0.0 -or $d -ge 1.0) {
        Write-Host "FAIL: NextDouble out of [0.0, 1.0) range: $d"
        exit 1
    }
}
Write-Host "PASS"
exit 0
