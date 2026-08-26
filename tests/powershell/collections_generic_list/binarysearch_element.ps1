# vybe-test: powershell/collections_generic_list/binarysearch_element
$list = [System.Collections.Generic.List[int]]::new([int[]]@(10, 20, 30, 40, 50))
$idx = $list.BinarySearch(30)
if ($idx -ne 2) {
    Write-Host "FAIL: BinarySearch expected index 2, got $idx"
    exit 1
}
Write-Host "PASS"
exit 0
