# vybe-test: powershell/math_floating_point_epsilon/epsilon_delta_comparison_pattern
$a = 0.1 + 0.2
$b = 0.3
$epsilon = 1e-9
$isEqual = [math]::Abs($a - $b) -lt $epsilon
if (-not $isEqual) {
    Write-Host "FAIL: Delta comparison with epsilon failed"
    exit 1
}
Write-Host "PASS"
exit 0
