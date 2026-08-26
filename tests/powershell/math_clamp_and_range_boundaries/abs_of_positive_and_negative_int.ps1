# vybe-test: powershell/math_clamp_and_range_boundaries/abs_of_positive_and_negative_int
$a1 = [math]::Abs(100)
$a2 = [math]::Abs(-100)
if ($a1 -ne 100 -or $a2 -ne 100) {
    Write-Host "FAIL: Abs int failed"
    exit 1
}
Write-Host "PASS"
exit 0
