# vybe-test: powershell/type_memory_extensions_and_spans/memory_copyto_target_memory
$src = [System.Memory[int]]::new([int[]]@(10, 20, 30))
$dstArr = [int[]]::new(3)
$dst = [System.Memory[int]]::new($dstArr)
$src.CopyTo($dst)
if ($dstArr[0] -ne 10 -or $dstArr[2] -ne 30) { Write-Host "FAIL: Memory CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
