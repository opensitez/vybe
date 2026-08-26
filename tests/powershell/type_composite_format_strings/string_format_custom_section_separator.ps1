# vybe-test: powershell/type_composite_format_strings/string_format_custom_section_separator
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:#;(#);zero}", -5)
if ($res -ne "(5)") { Write-Host "FAIL: Custom section format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
