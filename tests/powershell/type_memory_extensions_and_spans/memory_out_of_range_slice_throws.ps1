# vybe-test: powershell/type_memory_extensions_and_spans/memory_out_of_range_slice_throws
$mem = [System.MemoryExtensions]::AsMemory("Short")
$caught = $false
try {
    $x = $mem.Slice(10, 5)
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) { Write-Host "FAIL: ArgumentOutOfRangeException expected on out-of-range slice"; exit 1 }
Write-Host "PASS"; exit 0
