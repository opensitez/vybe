# vybe-test: powershell/math_floating_point_epsilon/bitdecrement_zero_returns_negative_epsilon
$prev = [math]::BitDecrement(0.0)
if ($prev -ne -[double]::Epsilon) {
    Write-Host "FAIL: BitDecrement(0.0) expected -Epsilon, got $prev"
    exit 1
}
Write-Host "PASS"
exit 0
