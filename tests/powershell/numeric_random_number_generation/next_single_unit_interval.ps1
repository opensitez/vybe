# vybe-test: powershell/numeric_random_number_generation/next_single_unit_interval
$rnd = [System.Random]::new()
$s = $rnd.NextSingle()
if ($s -lt 0.0 -or $s -ge 1.0) {
    Write-Host "FAIL: NextSingle out of range: $s"
    exit 1
}
Write-Host "PASS"
exit 0
