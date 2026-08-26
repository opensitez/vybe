# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_toarray_top_to_bottom
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
$cs.Push(1); $cs.Push(2); $cs.Push(3)
$arr = $cs.ToArray()
if ($arr[0] -ne 3 -or $arr[2] -ne 1) { Write-Host "FAIL: ToArray top to bottom failed"; exit 1 }
Write-Host "PASS"; exit 0
