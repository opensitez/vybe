# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_in_foreach_loop
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new([int[]]@(10, 20, 30))
$sum = 0
foreach ($item in $cs) { $sum += $item }
if ($sum -ne 60) { Write-Host "FAIL: foreach on ConcurrentStack failed"; exit 1 }
Write-Host "PASS"; exit 0
