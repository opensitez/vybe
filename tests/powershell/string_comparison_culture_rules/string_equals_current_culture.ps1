# vybe-test: powershell/string_comparison_culture_rules/string_equals_current_culture
$s1 = "abc"; $s2 = "ABC"
$eq = [string]::Equals($s1, $s2, [System.StringComparison]::CurrentCultureIgnoreCase)
if (-not $eq) { Write-Host "FAIL: CurrentCultureIgnoreCase failed"; exit 1 }
Write-Host "PASS"; exit 0
