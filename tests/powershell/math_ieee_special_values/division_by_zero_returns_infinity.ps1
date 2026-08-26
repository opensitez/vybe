# vybe-test: powershell/math_ieee_special_values/division_by_zero_returns_infinity
$res = 1.0 / 0.0
if (-not [double]::IsPositiveInfinity($res)) {
    Write-Host "FAIL: 1.0 / 0.0 should produce PositiveInfinity, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
