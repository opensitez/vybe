# vybe-test: powershell/math_floating_point_epsilon/double_parse_very_small_exponent
$val = [double]::Parse("1.0e-300")
if ($val -le 0.0) {
    Write-Host "FAIL: Parse 1e-300 failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
