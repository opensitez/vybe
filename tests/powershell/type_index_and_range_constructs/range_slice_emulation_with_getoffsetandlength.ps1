# vybe-test: powershell/type_index_and_range_constructs/range_slice_emulation_with_getoffsetandlength
$arr = [int[]]@(10, 20, 30, 40, 50)
$range = [System.Range]::new([System.Index]::FromStart(1), [System.Index]::FromEnd(1))
$t = $range.GetOffsetAndLength($arr.Length)
$offset = if ($t.Item1 -ne $null) { $t.Item1 } else { $t.Offset }
$len = if ($t.Item2 -ne $null) { $t.Item2 } else { $t.Length }
$slice = [int[]]::new($len)
[System.Array]::Copy($arr, $offset, $slice, 0, $len)
if ($slice.Length -ne 3 -or $slice[0] -ne 20 -or $slice[2] -ne 40) {
    Write-Host "FAIL: Range slice emulation failed"
    exit 1
}
Write-Host "PASS"
exit 0
