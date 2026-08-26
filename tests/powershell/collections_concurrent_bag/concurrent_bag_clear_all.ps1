# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_clear_all
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
$bag.Add(1); $bag.Add(2)
$bag.Clear()
if ($bag.Count -ne 0 -or -not $bag.IsEmpty) { Write-Host "FAIL: ConcurrentBag Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
