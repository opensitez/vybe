# vybe-test: powershell/string_comparison_culture_rules/string_comparison_mode_check_13
$c = [System.StringComparison]::Ordinal
$res = [string]::Compare("item_13", "ITEM_13", [System.StringComparison]::OrdinalIgnoreCase)
if ($res -ne 0) { Write-Host "FAIL: Comparison check failed"; exit 1 }
Write-Host "PASS"; exit 0
