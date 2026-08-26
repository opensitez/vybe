# vybe-test: powershell/collections_generic_hashset/removewhere_predicate
$set = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2, 3, 4, 5, 6))
$removedCount = $set.RemoveWhere([System.Predicate[int]]{ param($x) $x % 2 -eq 0 })
if ($removedCount -ne 3 -or $set.Count -ne 3 -or $set.Contains(2)) {
    Write-Host "FAIL: RemoveWhere predicate failed"
    exit 1
}
Write-Host "PASS"
exit 0
