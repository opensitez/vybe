# vybe-test: powershell/type_index_and_range_constructs/index_hashcode_consistency
$i1 = [System.Index]::FromEnd(3)
$i2 = [System.Index]::FromEnd(3)
if ($i1.GetHashCode() -ne $i2.GetHashCode()) { Write-Host "FAIL: Index HashCode failed"; exit 1 }
Write-Host "PASS"; exit 0
