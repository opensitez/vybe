# vybe-test: powershell/math_decimal_high_precision/decimal_negation
[decimal]$d = 45.67
$neg = -$d
if ($neg -ne [decimal]-45.67) {
    Write-Host "FAIL: Decimal negation failed, got $neg"
    exit 1
}
Write-Host "PASS"
exit 0
