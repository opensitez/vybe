# vybe-test: powershell/generic_types/generic_list_find_all
$list = [System.Collections.Generic.List[int]]::new()
$list.AddRange([int[]]@(1, 5, 10, 15, 20))
$match = $list.FindAll([Predicate[int]]{ param($x) $x -gt 9 })
if ($match.Count -ne 3) {
    Write-Host "FAIL: FindAll > 9 expected 3 items, got $($match.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
