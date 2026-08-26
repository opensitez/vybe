# vybe-test: powershell/math_rounding_midpoint_modes/decimal_type_rounding_with_mode
[decimal]$d = 2.555
$r = [math]::Round($d, 2, [System.MidpointRounding]::AwayFromZero)
if ($r -ne 2.56) {
    Write-Host "FAIL: Decimal rounding with mode failed, got $r"
    exit 1
}
Write-Host "PASS"
exit 0
