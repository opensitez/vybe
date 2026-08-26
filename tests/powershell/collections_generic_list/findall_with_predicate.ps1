# vybe-test: powershell/collections_generic_list/findall_with_predicate
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3, 4, 5, 6))
$evenList = $list.FindAll([System.Predicate[int]]{ param($x) $x % 2 -eq 0 })
if ($evenList.Count -ne 3 -or $evenList[0] -ne 2 -or $evenList[2] -ne 6) {
    Write-Host "FAIL: FindAll predicate failed"
    exit 1
}
Write-Host "PASS"
exit 0
