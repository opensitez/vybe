# vybe-test: powershell/type_index_and_range_constructs/index_equality_same_value_and_direction
$i1 = [System.Index]::FromStart(4)
$i2 = [System.Index]::FromStart(4)
$i3 = [System.Index]::FromEnd(4)
if (-not $i1.Equals($i2) -or $i1.Equals($i3)) { Write-Host "FAIL: Index Equals failed"; exit 1 }
Write-Host "PASS"; exit 0
