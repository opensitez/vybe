# vybe-test: powershell/collections_generic_hashset/construct_from_enumerable_array
$orig = @(1, 1, 2, 2, 3, 3, 4)
$set = [System.Collections.Generic.HashSet[int]]::new([int[]]$orig)
if ($set.Count -ne 4) {
    Write-Host "FAIL: HashSet constructor deduping failed"
    exit 1
}
Write-Host "PASS"
exit 0
