# vybe-test: powershell/math_ieee_special_values/neg_infinity_less_than_minvalue
$ninf = [double]::NegativeInfinity
$min = [double]::MinValue
if (-not ($ninf -lt $min)) {
    Write-Host "FAIL: NegativeInfinity should be less than Double.MinValue"
    exit 1
}
Write-Host "PASS"
exit 0
