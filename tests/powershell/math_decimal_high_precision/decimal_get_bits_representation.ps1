# vybe-test: powershell/math_decimal_high_precision/decimal_get_bits_representation
[decimal]$d = 100
$bits = [decimal]::GetBits($d)
if ($bits.Length -ne 4 -or $bits[0] -ne 100) {
    Write-Host "FAIL: Decimal GetBits failed"
    exit 1
}
Write-Host "PASS"
exit 0
