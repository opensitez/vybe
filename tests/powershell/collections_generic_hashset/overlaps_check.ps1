# vybe-test: powershell/collections_generic_hashset/overlaps_check
$s1 = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2, 3))
$s2 = [System.Collections.Generic.HashSet[int]]::new([int[]]@(3, 4, 5))
$s3 = [System.Collections.Generic.HashSet[int]]::new([int[]]@(6, 7, 8))
if (-not $s1.Overlaps($s2) -or $s1.Overlaps($s3)) {
    Write-Host "FAIL: Overlaps check failed"
    exit 1
}
Write-Host "PASS"
exit 0
