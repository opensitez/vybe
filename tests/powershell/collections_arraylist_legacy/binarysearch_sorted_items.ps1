# vybe-test: powershell/collections_arraylist_legacy/binarysearch_sorted_items
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(10, 20, 30, 40, 50))
$idx = $al.BinarySearch(40)
if ($idx -ne 3) {
    Write-Host "FAIL: BinarySearch expected 3, got $idx"
    exit 1
}
Write-Host "PASS"
exit 0
