# vybe-test: powershell/type_composite_format_strings/string_format_hexadecimal_lowercase_and_uppercase
$hexU = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:X4}", 255)
$hexL = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:x4}", 255)
if ($hexU -ne "00FF" -or $hexL -ne "00ff") { Write-Host "FAIL: Hex format failed, got $hexU, $hexL"; exit 1 }
Write-Host "PASS"; exit 0
