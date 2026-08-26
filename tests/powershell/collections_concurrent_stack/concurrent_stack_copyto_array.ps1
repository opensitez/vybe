# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_copyto_array
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new([int[]]@(1, 2))
[int[]]$arr = [int[]]::new(2)
$cs.CopyTo($arr, 0)
if ($arr.Length -ne 2) { Write-Host "FAIL: CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
