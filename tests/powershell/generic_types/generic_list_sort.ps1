# vybe-test: powershell/generic_types/generic_list_sort
$list = [System.Collections.Generic.List[int]]::new()
$list.AddRange([int[]]@(30, 10, 20))
$list.Sort()
if ($list[0] -ne 10 -or $list[1] -ne 20 -or $list[2] -ne 30) {
    Write-Host "FAIL: List.Sort() expected 10, 20, 30"
    exit 1
}
Write-Host "PASS"
exit 0
