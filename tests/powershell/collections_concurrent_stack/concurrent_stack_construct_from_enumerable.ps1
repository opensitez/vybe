# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_construct_from_enumerable
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new([int[]]@(1, 2, 3))
if ($cs.Count -ne 3) { Write-Host "FAIL: Constructor from enumerable failed"; exit 1 }
Write-Host "PASS"; exit 0
