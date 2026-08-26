# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_multiple_takes_until_empty
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new([int[]]@(1, 2, 3))
[int]$outVal = 0
$taken = 0
while ($bag.TryTake([ref]$outVal)) { $taken++ }
if ($taken -ne 3 -or -not $bag.IsEmpty) { Write-Host "FAIL: Multiple takes until empty failed"; exit 1 }
Write-Host "PASS"; exit 0
