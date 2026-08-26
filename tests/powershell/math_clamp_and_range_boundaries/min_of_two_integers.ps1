# vybe-test: powershell/math_clamp_and_range_boundaries/min_of_two_integers
$m = [math]::Min(42, 17)
if ($m -ne 17) {
    Write-Host "FAIL: Min expected 17, got $m"
    exit 1
}
Write-Host "PASS"
exit 0
