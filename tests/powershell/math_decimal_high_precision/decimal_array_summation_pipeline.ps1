# vybe-test: powershell/math_decimal_high_precision/decimal_array_summation_pipeline
[decimal[]]$items = @([decimal]10.25, [decimal]20.50, [decimal]30.25)
$sum = [decimal]0
foreach ($it in $items) { $sum += $it }
if ($sum -ne [decimal]61.00) {
    Write-Host "FAIL: Decimal summation pipeline failed, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
