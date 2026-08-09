# vybe-test: powershell/generic_types/generic_list_binary_search
$list = [System.Collections.Generic.List[int]]::new()
$list.AddRange([int[]]@(10, 20, 30, 40, 50))
$idx = $list.BinarySearch(30)
if ($idx -ne 2) {
    Write-Host "FAIL: BinarySearch(30) expected index 2, got $idx"
    exit 1
}
Write-Host "PASS"
exit 0
