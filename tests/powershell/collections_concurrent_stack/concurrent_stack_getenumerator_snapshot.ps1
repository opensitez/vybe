# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_getenumerator_snapshot
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new([int[]]@(1, 2))
$enum = $cs.GetEnumerator()
$cs.Push(3)
$list = [System.Collections.Generic.List[int]]::new()
while ($enum.MoveNext()) { $list.Add($enum.Current) }
if ($list.Count -lt 2) { Write-Host "FAIL: Enumerator snapshot failed"; exit 1 }
Write-Host "PASS"; exit 0
