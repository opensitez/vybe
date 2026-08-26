# vybe-test: powershell/math_floating_point_epsilon/double_max_value_positive
$max = [double]::MaxValue
if ($max -le 0.0 -or $max -lt 1.79e308) {
    Write-Host "FAIL: Double.MaxValue mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
