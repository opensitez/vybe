# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_decimal_precision
[decimal]$d = 99.99
$val = [math]::Clamp($d, [decimal]0, [decimal]50)
if ($val -ne 50) {
    Write-Host "FAIL: Clamp decimal failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
