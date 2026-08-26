# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_measure_object_sum
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new([int[]]@(10, 20, 30))
$m = $cs | Measure-Object -Sum
if ($m.Sum -ne 60) { Write-Host "FAIL: Measure-Object on ConcurrentStack failed"; exit 1 }
Write-Host "PASS"; exit 0
