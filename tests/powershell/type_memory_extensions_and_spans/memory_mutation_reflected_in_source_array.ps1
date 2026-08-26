# vybe-test: powershell/type_memory_extensions_and_spans/memory_mutation_reflected_in_source_array
$arr = [int[]]@(1, 2, 3)
$mem = [System.Memory[int]]::new($arr)
$arr[1] = 99
$recovered = $mem.ToArray()
if ($recovered[1] -ne 99) {
    Write-Host "FAIL: Array mutation not reflected in Memory"
    exit 1
}
Write-Host "PASS"
exit 0
