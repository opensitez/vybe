# vybe-test: powershell/collections_generic_list/reverse_in_place
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3))
$list.Reverse()
if ($list[0] -ne 3 -or $list[1] -ne 2 -or $list[2] -ne 1) {
    Write-Host "FAIL: Reverse failed"
    exit 1
}
Write-Host "PASS"
exit 0
