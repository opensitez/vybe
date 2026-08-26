# vybe-test: powershell/math_ieee_special_values/isnegative_check_negative_zero
$posZero = 0.0
$negZero = -0.0
if ([double]::IsNegative($posZero) -or -not [double]::IsNegative($negZero)) {
    Write-Host "FAIL: IsNegative on negative zero failed"
    exit 1
}
Write-Host "PASS"
exit 0
