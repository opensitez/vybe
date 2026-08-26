# vybe-test: powershell/type_composite_format_strings/string_format_enum_format_specifier
$d = [System.DayOfWeek]::Wednesday
$resD = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:D}", $d)
$resG = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:G}", $d)
if ($resD -ne "3" -or $resG -ne "Wednesday") { Write-Host "FAIL: Enum format specifier failed, got $resD, $resG"; exit 1 }
Write-Host "PASS"; exit 0
