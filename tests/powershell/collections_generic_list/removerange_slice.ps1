# vybe-test: powershell/collections_generic_list/removerange_slice
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3, 4, 5))
$list.RemoveRange(1, 3) # remove 2, 3, 4
if ($list.Count -ne 2 -or $list[0] -ne 1 -or $list[1] -ne 5) {
    Write-Host "FAIL: RemoveRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
