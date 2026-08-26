# vybe-test: powershell/math_decimal_high_precision/overflow_exception_on_decimal_multiply
$caught = $false
try {
    $x = [decimal]::MaxValue * 2
} catch [System.OverflowException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected OverflowException on Decimal.MaxValue * 2"
    exit 1
}
Write-Host "PASS"
exit 0
