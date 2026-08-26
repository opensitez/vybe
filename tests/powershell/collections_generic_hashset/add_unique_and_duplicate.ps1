# vybe-test: powershell/collections_generic_hashset/add_unique_and_duplicate
$set = [System.Collections.Generic.HashSet[int]]::new()
$added1 = $set.Add(10)
$added2 = $set.Add(20)
$addedDup = $set.Add(10)
if (-not $added1 -or -not $added2 -or $addedDup -or $set.Count -ne 2) {
    Write-Host "FAIL: HashSet unique/duplicate Add failed"
    exit 1
}
Write-Host "PASS"
exit 0
