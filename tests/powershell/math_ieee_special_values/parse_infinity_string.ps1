# vybe-test: powershell/math_ieee_special_values/parse_infinity_string
$posInf = [double]::PositiveInfinity
$isInf = [double]::IsPositiveInfinity($posInf)
if (-not $isInf) {
    Write-Host "FAIL: Positive Infinity check failed"
    exit 1
}
Write-Host "PASS"
exit 0
