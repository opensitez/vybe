# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_construct_from_enumerable
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new([int[]]@(1, 2, 3, 4, 5))
if ($bag.Count -ne 5) { Write-Host "FAIL: Constructor from enumerable failed"; exit 1 }
Write-Host "PASS"; exit 0
