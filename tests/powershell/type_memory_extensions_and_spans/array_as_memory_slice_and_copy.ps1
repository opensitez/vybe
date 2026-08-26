# vybe-test: powershell/type_memory_extensions_and_spans/array_as_memory_slice_and_copy
$arr = [int[]]@(10, 20, 30, 40, 50)
$mem = [System.MemoryExtensions]::AsMemory($arr, 1, 3)
$slice = $mem.ToArray()
if ($slice.Length -ne 3 -or $slice[0] -ne 20 -or $slice[2] -ne 40) { Write-Host "FAIL: Array AsMemory failed"; exit 1 }
Write-Host "PASS"; exit 0
