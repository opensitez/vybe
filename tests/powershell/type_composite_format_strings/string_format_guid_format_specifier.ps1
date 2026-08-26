# vybe-test: powershell/type_composite_format_strings/string_format_guid_format_specifier
$g = [guid]::Parse("12345678-1234-1234-1234-123456789abc")
$resN = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:N}", $g)
$resB = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:B}", $g)
if ($resN -ne "12345678123412341234123456789abc" -or -not $resB.StartsWith("{")) { Write-Host "FAIL: Guid format failed, got $resN, $resB"; exit 1 }
Write-Host "PASS"; exit 0
