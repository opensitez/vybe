# vybe-test: powershell/type_composite_format_strings/string_format_date_custom_specifiers
$dt = [datetime]::ParseExact("2026-08-26", "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:yyyy/MM/dd}", $dt)
if ($res -ne "2026/08/26") { Write-Host "FAIL: Date composite format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
