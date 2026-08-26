# vybe-test: powershell/math_ieee_special_values/positive_zero_equals_negative_zero
$posZero = 0.0
$negZero = -0.0
if ($posZero -ne $negZero) {
    Write-Host "FAIL: 0.0 and -0.0 must compare equal"
    exit 1
}
Write-Host "PASS"
exit 0
