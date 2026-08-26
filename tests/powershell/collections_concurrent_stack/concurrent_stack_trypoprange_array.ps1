# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_trypoprange_array
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
$cs.PushRange([int[]]@(1, 2, 3, 4, 5))
[int[]]$buffer = [int[]]::new(3)
$popped = $cs.TryPopRange($buffer)
if ($popped -ne 3 -or $cs.Count -ne 2) { Write-Host "FAIL: TryPopRange failed"; exit 1 }
Write-Host "PASS"; exit 0
