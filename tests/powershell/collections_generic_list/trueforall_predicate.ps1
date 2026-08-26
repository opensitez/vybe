# vybe-test: powershell/collections_generic_list/trueforall_predicate
$list = [System.Collections.Generic.List[int]]::new([int[]]@(2, 4, 6, 8))
$allEven = $list.TrueForAll([System.Predicate[int]]{ param($x) $x % 2 -eq 0 })
if (-not $allEven) {
    Write-Host "FAIL: TrueForAll expected true"
    exit 1
}
Write-Host "PASS"
exit 0
