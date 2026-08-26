# vybe-test: powershell/type_index_and_range_constructs/range_equality_same_boundaries
$r1 = [System.Range]::new([System.Index]::FromStart(1), [System.Index]::FromStart(5))
$r2 = [System.Range]::new([System.Index]::FromStart(1), [System.Index]::FromStart(5))
if (-not $r1.Equals($r2)) { Write-Host "FAIL: Range Equals failed"; exit 1 }
Write-Host "PASS"; exit 0
