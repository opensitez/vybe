# vybe-test: powershell/math_decimal_high_precision/math_round_decimal_away_from_zero
[decimal]$d = 1.005
$r = [math]::Round($d, 2, [System.MidpointRounding]::AwayFromZero)
if ($r -ne [decimal]1.01) {
    Write-Host "FAIL: Decimal Round AwayFromZero failed, got $r"
    exit 1
}
Write-Host "PASS"
exit 0
