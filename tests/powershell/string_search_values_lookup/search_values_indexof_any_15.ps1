# vybe-test: powershell/string_search_values_lookup/search_values_indexof_any_15
$chars = [char[]]@('a', 'e', 'i', 'o', 'u')
$str = "xyz_test_15_alpha"
$idx = $str.IndexOfAny($chars)
if ($idx -lt 0) { Write-Host "FAIL: IndexOfAny failed"; exit 1 }
Write-Host "PASS"; exit 0
