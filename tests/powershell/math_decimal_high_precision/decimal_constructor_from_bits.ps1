# vybe-test: powershell/math_decimal_high_precision/decimal_constructor_from_bits
[int[]]$bits = @(255, 0, 0, 0)
$d = [decimal]::new($bits)
if ($d -ne [decimal]255) {
    Write-Host "FAIL: Decimal constructor from bits failed, got $d"
    exit 1
}
Write-Host "PASS"
exit 0
