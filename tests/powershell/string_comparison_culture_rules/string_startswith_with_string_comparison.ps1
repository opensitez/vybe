# vybe-test: powershell/string_comparison_culture_rules/string_startswith_with_string_comparison
$str = "HTTPS://example.com"
$sw = $str.StartsWith("https://", [System.StringComparison]::OrdinalIgnoreCase)
if (-not $sw) { Write-Host "FAIL: StartsWith with StringComparison failed"; exit 1 }
Write-Host "PASS"; exit 0
