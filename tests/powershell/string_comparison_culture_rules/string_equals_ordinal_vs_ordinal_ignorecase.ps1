# vybe-test: powershell/string_comparison_culture_rules/string_equals_ordinal_vs_ordinal_ignorecase
$s1 = "hello"; $s2 = "HELLO"
$eqOrd = [string]::Equals($s1, $s2, [System.StringComparison]::Ordinal)
$eqOrdIg = [string]::Equals($s1, $s2, [System.StringComparison]::OrdinalIgnoreCase)
if ($eqOrd -or -not $eqOrdIg) { Write-Host "FAIL: StringComparison Ordinal vs OrdinalIgnoreCase failed"; exit 1 }
Write-Host "PASS"; exit 0
