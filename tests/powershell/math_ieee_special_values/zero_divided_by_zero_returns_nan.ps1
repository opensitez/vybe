# vybe-test: powershell/math_ieee_special_values/zero_divided_by_zero_returns_nan
$res = 0.0 / 0.0
if (-not [double]::IsNaN($res)) {
    Write-Host "FAIL: 0.0 / 0.0 should produce NaN, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
