# vybe-test: powershell/math_ieee_special_values/infinity_minus_infinity_returns_nan
$p = [double]::PositiveInfinity
$res = $p - $p
if (-not [double]::IsNaN($res)) {
    Write-Host "FAIL: Inf - Inf should produce NaN, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
