# vybe-test: powershell/type_index_and_range_constructs/range_getoffsetandlength_calculation
$range = [System.Range]::new([System.Index]::FromStart(2), [System.Index]::FromEnd(2))
$t = $range.GetOffsetAndLength(10)
$offset = if ($t.Item1 -ne $null) { $t.Item1 } else { $t.Offset }
$len = if ($t.Item2 -ne $null) { $t.Item2 } else { $t.Length }
if ($offset -ne 2 -or $len -ne 6) {
    Write-Host "FAIL: Range GetOffsetAndLength failed, got $offset, $len"
    exit 1
}
Write-Host "PASS"
exit 0
