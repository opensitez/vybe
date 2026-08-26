# vybe-test: powershell/type_memory_extensions_and_spans/memory_trycopyto_insufficient_space_returns_false
$src = [System.Memory[int]]::new([int[]]@(1, 2, 3, 4))
$dst = [System.Memory[int]]::new([int[]]::new(2))
$ok = $src.TryCopyTo($dst)
if ($ok) { Write-Host "FAIL: TryCopyTo should return false on small destination"; exit 1 }
Write-Host "PASS"; exit 0
