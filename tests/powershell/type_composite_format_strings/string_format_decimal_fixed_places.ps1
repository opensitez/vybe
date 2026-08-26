# vybe-test: powershell/type_composite_format_strings/string_format_decimal_fixed_places
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:F2}", 12.3456)
if ($res -ne "12.35") { Write-Host "FAIL: Fixed point decimal format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
