# vybe-test: powershell/string_comparison_culture_rules/string_equals_invariant_culture
$s1 = "test"; $s2 = "TEST"
$eq = [string]::Equals($s1, $s2, [System.StringComparison]::InvariantCultureIgnoreCase)
if (-not $eq) { Write-Host "FAIL: InvariantCultureIgnoreCase failed"; exit 1 }
Write-Host "PASS"; exit 0
