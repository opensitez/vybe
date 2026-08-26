# vybe-test: powershell/math_floating_point_epsilon/bitincrement_zero_returns_epsilon
$next = [math]::BitIncrement(0.0)
if ($next -ne [double]::Epsilon) {
    Write-Host "FAIL: BitIncrement(0.0) expected Epsilon, got $next"
    exit 1
}
Write-Host "PASS"
exit 0
