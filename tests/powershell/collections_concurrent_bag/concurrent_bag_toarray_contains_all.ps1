# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_toarray_contains_all
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
$bag.Add(1); $bag.Add(2); $bag.Add(3)
$arr = $bag.ToArray()
if ($arr.Length -ne 3 -or -not ($arr -contains 1) -or -not ($arr -contains 3)) { Write-Host "FAIL: ToArray failed"; exit 1 }
Write-Host "PASS"; exit 0
