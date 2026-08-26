# vybe-test: powershell/math_decimal_high_precision/decimal_ceiling_and_floor
[decimal]$d = 12.34
$c = [decimal]::Ceiling($d)
$f = [decimal]::Floor($d)
if ($c -ne [decimal]13 -or $f -ne [decimal]12) {
    Write-Host "FAIL: Decimal Ceiling/Floor failed"
    exit 1
}
Write-Host "PASS"
exit 0
