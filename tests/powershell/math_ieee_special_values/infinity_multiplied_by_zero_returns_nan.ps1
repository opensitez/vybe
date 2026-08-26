# vybe-test: powershell/math_ieee_special_values/infinity_multiplied_by_zero_returns_nan
$p = [double]::PositiveInfinity
$res = $p * 0.0
if (-not [double]::IsNaN($res)) {
    Write-Host "FAIL: Inf * 0 should produce NaN, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
