# vybe-test: powershell/type_memory_extensions_and_spans/memory_byte_equality_comparison
$bytes = [byte[]]@(1, 2, 3)
$b1 = [System.ReadOnlyMemory[byte]]::new($bytes)
$b2 = [System.ReadOnlyMemory[byte]]::new($bytes)
if (-not $b1.Equals($b2)) {
    Write-Host "FAIL: ReadOnlyMemory[byte] Equals failed"
    exit 1
}
Write-Host "PASS"
exit 0
