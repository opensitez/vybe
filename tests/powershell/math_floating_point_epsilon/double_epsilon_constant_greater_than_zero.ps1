# vybe-test: powershell/math_floating_point_epsilon/double_epsilon_constant_greater_than_zero
$eps = [double]::Epsilon
if ($eps -le 0.0) {
    Write-Host "FAIL: Double.Epsilon must be greater than zero, got $eps"
    exit 1
}
Write-Host "PASS"
exit 0
