# vybe-test: powershell/string_comparison_culture_rules/string_compare_ordinal_returns_integer_difference
$diff = [string]::Compare("a", "c", [System.StringComparison]::Ordinal)
if ($diff -ge 0) { Write-Host "FAIL: String Compare expected negative difference"; exit 1 }
Write-Host "PASS"; exit 0
