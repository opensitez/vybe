# vybe-test: powershell/type_composite_format_strings/string_format_alignment_padding_right
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0,10}", "Test")
if ($res.Length -ne 10 -or -not $res.StartsWith("      Test")) { Write-Host "FAIL: Right alignment padding failed, got '$res'"; exit 1 }
Write-Host "PASS"; exit 0
