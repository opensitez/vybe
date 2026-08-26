# vybe-test: powershell/math_floating_point_epsilon/single_epsilon_constant_greater_than_zero
$eps = [float]::Epsilon
if ($eps -le 0.0) {
    Write-Host "FAIL: Float.Epsilon must be greater than zero, got $eps"
    exit 1
}
Write-Host "PASS"
exit 0
