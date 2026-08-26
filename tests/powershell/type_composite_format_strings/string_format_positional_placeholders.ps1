# vybe-test: powershell/type_composite_format_strings/string_format_positional_placeholders
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0} is {1}", "Vybe", "Fast")
if ($res -ne "Vybe is Fast") { Write-Host "FAIL: string.Format positional failed, got $res"; exit 1 }
Write-Host "PASS"; exit 0
