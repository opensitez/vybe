# vybe-test: powershell/type_index_and_range_constructs/index_getoffset_calculation
$idx = [System.Index]::FromEnd(2)
$offset = $idx.GetOffset(10)
if ($offset -ne 8) { Write-Host "FAIL: Index GetOffset expected 8, got $offset"; exit 1 }
Write-Host "PASS"; exit 0
