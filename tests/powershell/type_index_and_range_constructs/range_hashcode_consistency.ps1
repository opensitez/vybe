# vybe-test: powershell/type_index_and_range_constructs/range_hashcode_consistency
$r1 = [System.Range]::new([System.Index]::FromStart(2), [System.Index]::FromStart(8))
$r2 = [System.Range]::new([System.Index]::FromStart(2), [System.Index]::FromStart(8))
if ($r1.GetHashCode() -ne $r2.GetHashCode()) { Write-Host "FAIL: Range HashCode failed"; exit 1 }
Write-Host "PASS"; exit 0
