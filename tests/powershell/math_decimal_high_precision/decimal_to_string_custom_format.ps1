# vybe-test: powershell/math_decimal_high_precision/decimal_to_string_custom_format
[decimal]$d = 1234.5
$str = $d.ToString("C2", [System.Globalization.CultureInfo]::InvariantCulture)
if (-not $str.Contains("1,234.50")) {
    Write-Host "FAIL: Decimal currency formatting failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
