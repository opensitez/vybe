# vybe-test: powershell/collections_sorted_set/sorted_set_removewhere_predicate
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(1..6))
$rem = $ss.RemoveWhere([System.Predicate[int]]{ param($x) $x % 2 -eq 0 })
if ($rem -ne 3 -or $ss.Count -ne 3 -or $ss.Contains(2)) { Write-Host "FAIL: RemoveWhere failed"; exit 1 }
Write-Host "PASS"; exit 0
