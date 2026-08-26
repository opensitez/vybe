# vybe-test: powershell/math_decimal_high_precision/divide_by_zero_exception_on_decimal
$caught = $false
try {
    $x = [decimal]10 / [decimal]0
} catch [System.DivideByZeroException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected DivideByZeroException on decimal division"
    exit 1
}
Write-Host "PASS"
exit 0
