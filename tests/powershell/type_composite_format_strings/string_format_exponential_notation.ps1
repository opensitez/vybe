# vybe-test: powershell/type_composite_format_strings/string_format_exponential_notation
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:E2}", 10500.0)
if (-not $res.StartsWith("1.05E+004") -and -not $res.StartsWith("1.05E+04")) { Write-Host "FAIL: Exponential format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
