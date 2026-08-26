# vybe-test: powershell/math_floating_point_epsilon/float_min_and_max_value
$min = [float]::MinValue
$max = [float]::MaxValue
if ($min -ge 0.0 -or $max -le 0.0) {
    Write-Host "FAIL: Float Min/Max failed"
    exit 1
}
Write-Host "PASS"
exit 0
