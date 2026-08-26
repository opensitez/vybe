# vybe-test: powershell/collections_generic_hashset/intersectwith_common_elements
$s1 = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2, 3, 4))
$s2 = [System.Collections.Generic.HashSet[int]]::new([int[]]@(3, 4, 5, 6))
$s1.IntersectWith($s2)
if ($s1.Count -ne 2 -or -not $s1.Contains(3) -or -not $s1.Contains(4)) {
    Write-Host "FAIL: IntersectWith failed"
    exit 1
}
Write-Host "PASS"
exit 0
