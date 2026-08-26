# vybe-test: powershell/type_memory_extensions_and_spans/string_as_memory_and_slice
$str = "Hello World"
$mem = [System.MemoryExtensions]::AsMemory($str, 0, 5)
if ($mem.Length -ne 5 -or $mem.ToString() -ne "Hello") { Write-Host "FAIL: AsMemory slice failed"; exit 1 }
Write-Host "PASS"; exit 0
