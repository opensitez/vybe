# vybe-test: powershell/type_index_and_range_constructs/range_tostring_format
$range = [System.Range]::new([System.Index]::FromStart(1), [System.Index]::FromEnd(1))
$str = $range.ToString()
if ($str -ne "1..^1") { Write-Host "FAIL: Range ToString expected '1..^1', got '$str'"; exit 1 }
Write-Host "PASS"; exit 0
