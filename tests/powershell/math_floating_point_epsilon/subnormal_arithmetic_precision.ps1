# vybe-test: powershell/math_floating_point_epsilon/subnormal_arithmetic_precision
$sub = [double]::Epsilon * 2.0
if ($sub -le 0.0 -or $sub -le [double]::Epsilon) {
    Write-Host "FAIL: Subnormal arithmetic failed"
    exit 1
}
Write-Host "PASS"
exit 0
