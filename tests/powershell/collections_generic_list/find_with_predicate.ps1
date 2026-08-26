# vybe-test: powershell/collections_generic_list/find_with_predicate
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 4, 7, 10, 13))
$found = $list.Find([System.Predicate[int]]{ param($x) $x -gt 8 })
if ($found -ne 10) {
    Write-Host "FAIL: Find predicate expected 10, got $found"
    exit 1
}
Write-Host "PASS"
exit 0
