# vybe-test: powershell/collections_generic_hashset/unionwith_combine_sets
$s1 = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2, 3))
$s2 = [System.Collections.Generic.HashSet[int]]::new([int[]]@(3, 4, 5))
$s1.UnionWith($s2)
if ($s1.Count -ne 5 -or -not $s1.Contains(1) -or -not $s1.Contains(5)) {
    Write-Host "FAIL: UnionWith failed"
    exit 1
}
Write-Host "PASS"
exit 0
