# vybe-test: powershell/collections_generic_hashset/issupersetof_true_and_false
$sub = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2))
$sup = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2, 3))
if (-not $sup.IsProperSupersetOf($sub) -or -not $sup.IsSupersetOf($sub)) {
    Write-Host "FAIL: IsSupersetOf failed"
    exit 1
}
Write-Host "PASS"
exit 0
