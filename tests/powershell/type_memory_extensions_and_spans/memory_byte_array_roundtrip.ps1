# vybe-test: powershell/type_memory_extensions_and_spans/memory_byte_array_roundtrip
$bytes = [byte[]]@(1, 2, 3, 4, 5)
$mem = [System.Memory[byte]]::new($bytes)
$arr = $mem.ToArray()
if ($arr.Length -ne 5 -or $arr[0] -ne 1 -or $arr[4] -ne 5) { Write-Host "FAIL: Memory[byte] roundtrip failed"; exit 1 }
Write-Host "PASS"; exit 0
