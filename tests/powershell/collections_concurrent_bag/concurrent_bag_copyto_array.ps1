# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_copyto_array
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
$bag.Add(5); $bag.Add(10)
[int[]]$arr = [int[]]::new(2)
$bag.CopyTo($arr, 0)
if ($arr.Length -ne 2) { Write-Host "FAIL: CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
