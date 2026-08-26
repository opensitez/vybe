# vybe-test: powershell/collections_sorted_set/sorted_set_add_and_contains
$ss = [System.Collections.Generic.SortedSet[int]]::new()
$ok1 = $ss.Add(20); $ok2 = $ss.Add(10); $ok3 = $ss.Add(20)
if (-not $ok1 -or -not $ok2 -or $ok3 -or -not $ss.Contains(10)) { Write-Host "FAIL: SortedSet Add/Contains failed"; exit 1 }
Write-Host "PASS"; exit 0
