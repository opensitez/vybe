# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_getenumerator_threadsafe_snapshot
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new([int[]]@(10, 20))
$enum = $bag.GetEnumerator()
$bag.Add(30)
$list = [System.Collections.Generic.List[int]]::new()
while ($enum.MoveNext()) { $list.Add($enum.Current) }
if ($list.Count -lt 2) { Write-Host "FAIL: Enumerator snapshot failed"; exit 1 }
Write-Host "PASS"; exit 0
