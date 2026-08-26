# vybe-test: powershell/math_ieee_special_values/isnegativeinfinity_check
$ninf = [double]::NegativeInfinity
if (-not [double]::IsNegativeInfinity($ninf) -or [double]::IsPositiveInfinity($ninf)) {
    Write-Host "FAIL: IsNegativeInfinity check failed"
    exit 1
}
Write-Host "PASS"
exit 0
