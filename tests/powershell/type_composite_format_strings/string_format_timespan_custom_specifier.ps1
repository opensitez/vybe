# vybe-test: powershell/type_composite_format_strings/string_format_timespan_custom_specifier
$ts = [timespan]::FromMinutes(90)
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:hh\:mm}", $ts)
if ($res -ne "01:30") { Write-Host "FAIL: TimeSpan format failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
