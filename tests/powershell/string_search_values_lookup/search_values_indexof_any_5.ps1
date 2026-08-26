# vybe-test: powershell/string_search_values_lookup/search_values_indexof_any_5
$chars = [char[]]@('a', 'e', 'i', 'o', 'u')
$str = "xyz_test_5_alpha"
$idx = $str.IndexOfAny($chars)
if ($idx -lt 0) { Write-Host "FAIL: IndexOfAny failed"; exit 1 }
Write-Host "PASS"; exit 0
