# vybe-test: powershell/type_composite_format_strings/string_format_currency_format_specifier
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:C0}", 100)
if (-not $res.Contains("100")) { Write-Host "FAIL: Currency format specifier failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
