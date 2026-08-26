# vybe-test: powershell/type_composite_format_strings/string_format_percent_specifier
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:P0}", 0.75)
if ($res -ne "75 %" -and $res -ne "75%") { Write-Host "FAIL: Percent format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
