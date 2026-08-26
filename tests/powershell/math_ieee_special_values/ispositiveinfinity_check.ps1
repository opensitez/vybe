# vybe-test: powershell/math_ieee_special_values/ispositiveinfinity_check
$inf = [double]::PositiveInfinity
if (-not [double]::IsPositiveInfinity($inf) -or [double]::IsNegativeInfinity($inf)) {
    Write-Host "FAIL: IsPositiveInfinity check failed"
    exit 1
}
Write-Host "PASS"
exit 0
