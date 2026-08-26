# vybe-test: powershell/type_composite_format_strings/string_format_escaped_curly_braces
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{{Value: {0}}}", 100)
if ($res -ne "{Value: 100}") { Write-Host "FAIL: Escaped curly braces failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
