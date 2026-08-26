# vybe-test: powershell/string_comparison_culture_rules/string_indexof_with_string_comparison_enum
$str = "The quick brown Fox jumps"
$idx = $str.IndexOf("fox", [System.StringComparison]::OrdinalIgnoreCase)
if ($idx -ne 16) { Write-Host "FAIL: IndexOf with StringComparison failed, got $idx"; exit 1 }
Write-Host "PASS"; exit 0
