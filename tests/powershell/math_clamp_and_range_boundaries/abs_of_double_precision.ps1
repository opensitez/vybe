# vybe-test: powershell/math_clamp_and_range_boundaries/abs_of_double_precision
$a = [math]::Abs(-123.456)
if ($a -ne 123.456) {
    Write-Host "FAIL: Abs double failed, got $a"
    exit 1
}
Write-Host "PASS"
exit 0
