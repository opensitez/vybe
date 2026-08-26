# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_clear_all
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
$cs.Push(1); $cs.Push(2)
$cs.Clear()
if ($cs.Count -ne 0 -or -not $cs.IsEmpty) { Write-Host "FAIL: Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
