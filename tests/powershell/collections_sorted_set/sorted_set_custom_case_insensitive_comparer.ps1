# vybe-test: powershell/collections_sorted_set/sorted_set_custom_case_insensitive_comparer
$comp = [System.StringComparer]::OrdinalIgnoreCase
$ss = [System.Collections.Generic.SortedSet[string]]::new($comp)
$ss.Add("HELLO")
if (-not $ss.Contains("hello") -or $ss.Add("hello")) { Write-Host "FAIL: Case insensitive SortedSet failed"; exit 1 }
Write-Host "PASS"; exit 0
