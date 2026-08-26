# vybe-test: powershell/math_decimal_high_precision/exact_financial_multiplication
[decimal]$price = 19.99
[decimal]$qty = 100
$total = $price * $qty
if ($total -ne [decimal]1999.00) {
    Write-Host "FAIL: Financial multiplication failed, got $total"
    exit 1
}
Write-Host "PASS"
exit 0
