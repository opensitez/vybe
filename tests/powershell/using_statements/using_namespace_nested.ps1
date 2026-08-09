# vybe-test: powershell/using_statements/using_namespace_nested
using namespace System.Collections.Concurrent

$dict = [ConcurrentDictionary[string, int]]::new()
[void]$dict.TryAdd("Key", 500)
if ($dict["Key"] -ne 500) {
    Write-Host "FAIL: ConcurrentDictionary short type lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
