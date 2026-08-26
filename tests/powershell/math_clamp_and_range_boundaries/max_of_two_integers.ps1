# vybe-test: powershell/math_clamp_and_range_boundaries/max_of_two_integers
$m = [math]::Max(42, 17)
if ($m -ne 42) {
    Write-Host "FAIL: Max expected 42, got $m"
    exit 1
}
Write-Host "PASS"
exit 0
