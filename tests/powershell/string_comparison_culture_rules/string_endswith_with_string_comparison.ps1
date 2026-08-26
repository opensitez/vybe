# vybe-test: powershell/string_comparison_culture_rules/string_endswith_with_string_comparison
$str = "image.PNG"
$ew = $str.EndsWith(".png", [System.StringComparison]::OrdinalIgnoreCase)
if (-not $ew) { Write-Host "FAIL: EndsWith with StringComparison failed"; exit 1 }
Write-Host "PASS"; exit 0
