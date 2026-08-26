# vybe-test: powershell/math_ieee_special_values/isnormal_and_issubnormal
$normal = 1.0
$subnormal = [double]::Epsilon
if (-not [double]::IsNormal($normal) -or -not [double]::IsSubnormal($subnormal)) {
    Write-Host "FAIL: IsNormal / IsSubnormal failed"
    exit 1
}
Write-Host "PASS"
exit 0
