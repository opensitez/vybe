# vybe-test: powershell/numeric_random_number_generation/next_equal_min_and_max_returns_min
$rnd = [System.Random]::new()
$v = $rnd.Next(5, 5)
if ($v -ne 5) {
    Write-Host "FAIL: Next(5, 5) expected 5, got $v"
    exit 1
}
Write-Host "PASS"
exit 0
