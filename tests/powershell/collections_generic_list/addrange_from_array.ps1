# vybe-test: powershell/collections_generic_list/addrange_from_array
$list = [System.Collections.Generic.List[int]]::new()
$list.AddRange([int[]]@(10, 20, 30, 40))
if ($list.Count -ne 4 -or $list[3] -ne 40) {
    Write-Host "FAIL: AddRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
