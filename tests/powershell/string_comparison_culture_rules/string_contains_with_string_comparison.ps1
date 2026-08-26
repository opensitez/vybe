# vybe-test: powershell/string_comparison_culture_rules/string_contains_with_string_comparison
$str = "PowerShell Core 7"
$has = $str.Contains("powershell", [System.StringComparison]::OrdinalIgnoreCase)
if (-not $has) { Write-Host "FAIL: Contains with StringComparison failed"; exit 1 }
Write-Host "PASS"; exit 0
