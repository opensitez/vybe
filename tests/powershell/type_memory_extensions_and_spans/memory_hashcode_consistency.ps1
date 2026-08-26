# vybe-test: powershell/type_memory_extensions_and_spans/memory_hashcode_consistency
$str = "Alpha"
$m1 = [System.MemoryExtensions]::AsMemory($str)
$m2 = [System.MemoryExtensions]::AsMemory($str)
if ($m1.GetHashCode() -ne $m2.GetHashCode()) {
    Write-Host "FAIL: Memory HashCode failed"
    exit 1
}
Write-Host "PASS"
exit 0
