# vybe-test: powershell/type_composite_format_strings/string_format_multiple_composite_arguments
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0} + {1} = {2}", 2, 3, 5)
if ($res -ne "2 + 3 = 5") { Write-Host "FAIL: Multiple composite arguments failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
