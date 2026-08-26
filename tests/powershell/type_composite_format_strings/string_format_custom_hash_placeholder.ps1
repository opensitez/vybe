# vybe-test: powershell/type_composite_format_strings/string_format_custom_hash_placeholder
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:###.##}", 123.456)
if ($res -ne "123.46") { Write-Host "FAIL: Custom hash format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
