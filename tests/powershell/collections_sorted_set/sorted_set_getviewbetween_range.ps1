# vybe-test: powershell/collections_sorted_set/sorted_set_getviewbetween_range
$ss = [System.Collections.Generic.SortedSet[int]]::new([int[]]@(10, 20, 30, 40, 50))
$view = $ss.GetViewBetween(20, 40)
if ($view.Count -ne 3 -or $view.Min -ne 20 -or $view.Max -ne 40) { Write-Host "FAIL: GetViewBetween failed"; exit 1 }
Write-Host "PASS"; exit 0
