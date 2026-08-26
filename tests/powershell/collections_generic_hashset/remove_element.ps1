# vybe-test: powershell/collections_generic_hashset/remove_element
$set = [System.Collections.Generic.HashSet[int]]::new([int[]]@(1, 2, 3))
$rem = $set.Remove(2)
if (-not $rem -or $set.Count -ne 2 -or $set.Contains(2)) {
    Write-Host "FAIL: HashSet Remove failed"
    exit 1
}
Write-Host "PASS"
exit 0
