# vybe-test: powershell/collections_generic_hashset/issubsetof_true_and_false
$sub = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2))
$sup = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2, 3, 4))
if (-not $sub.IsSubsetOf($sup) -or $sup.IsSubsetOf($sub)) {
    Write-Host "FAIL: IsSubsetOf failed"
    exit 1
}
Write-Host "PASS"
exit 0
