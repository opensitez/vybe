# vybe-test: powershell/math_ieee_special_values/infinity_greater_than_maxvalue
$inf = [double]::PositiveInfinity
$max = [double]::MaxValue
if (-not ($inf -gt $max)) {
    Write-Host "FAIL: PositiveInfinity should be greater than Double.MaxValue"
    exit 1
}
Write-Host "PASS"
exit 0
