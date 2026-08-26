# vybe-test: powershell/type_composite_format_strings/string_format_alignment_padding_left
$res = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0,-10}", "Test")
if ($res.Length -ne 10 -or -not $res.EndsWith("Test      ")) { Write-Host "FAIL: Left alignment padding failed, got '$res'"; exit 1 }
Write-Host "PASS"; exit 0
