# vybe-test: powershell/using_statements/using_namespace_concurrent
using namespace System.Collections.Concurrent

$bag = [ConcurrentBag[int]]::new()
$bag.Add(10)
$item = 0
$took = $bag.TryTake([ref]$item)
if (-not $took -or $item -ne 10) {
    Write-Host "FAIL: ConcurrentBag TryTake expected true and 10"
    exit 1
}
Write-Host "PASS"
exit 0
