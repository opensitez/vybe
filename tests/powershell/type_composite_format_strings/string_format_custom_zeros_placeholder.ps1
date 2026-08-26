# vybe-test: powershell/type_composite_format_strings/string_format_custom_zeros_placeholder
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:00000}", 42)
if ($res -ne "00042") { Write-Host "FAIL: Custom zeros format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
