# vybe-test: powershell/math_ieee_special_values/negative_division_by_zero_returns_neg_infinity
$res = -1.0 / 0.0
if (-not [double]::IsNegativeInfinity($res)) {
    Write-Host "FAIL: -1.0 / 0.0 should produce NegativeInfinity, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
