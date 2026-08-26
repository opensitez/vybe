# vybe-test: powershell/string_comparison_culture_rules/string_replace_with_string_comparison
$str = "Cat and DOG and cat"
$res = $str.Replace("cat", "bird", [System.StringComparison]::OrdinalIgnoreCase)
if ($res -ne "bird and DOG and bird") { Write-Host "FAIL: Replace with StringComparison failed, got '$res'"; exit 1 }
Write-Host "PASS"; exit 0
