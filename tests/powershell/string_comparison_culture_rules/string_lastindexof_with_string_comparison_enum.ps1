# vybe-test: powershell/string_comparison_culture_rules/string_lastindexof_with_string_comparison_enum
$str = "data DATA data DATA"
$idx = $str.LastIndexOf("data", [System.StringComparison]::OrdinalIgnoreCase)
if ($idx -ne 15) { Write-Host "FAIL: LastIndexOf with StringComparison failed, got $idx"; exit 1 }
Write-Host "PASS"; exit 0
