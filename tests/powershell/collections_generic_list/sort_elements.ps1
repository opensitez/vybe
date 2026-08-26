# vybe-test: powershell/collections_generic_list/sort_elements
$list = [System.Collections.Generic.List[int]]::new([int[]]@(5, 2, 8, 1))
$list.Sort()
if ($list[0] -ne 1 -or $list[1] -ne 2 -or $list[2] -ne 5 -or $list[3] -ne 8) {
    Write-Host "FAIL: Sort failed"
    exit 1
}
Write-Host "PASS"
exit 0
