# vybe-test: powershell/type_memory_extensions_and_spans/memory_equals_comparison
$str = "Test"
$m1 = [System.MemoryExtensions]::AsMemory($str)
$m2 = [System.MemoryExtensions]::AsMemory($str)
if (-not $m1.Equals($m2)) {
    Write-Host "FAIL: ReadOnlyMemory Equals failed"
    exit 1
}
Write-Host "PASS"
exit 0
