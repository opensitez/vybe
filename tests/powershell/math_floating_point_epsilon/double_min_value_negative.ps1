# vybe-test: powershell/math_floating_point_epsilon/double_min_value_negative
$min = [double]::MinValue
if ($min -ge 0.0 -or $min -gt -1.79e308) {
    Write-Host "FAIL: Double.MinValue mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
