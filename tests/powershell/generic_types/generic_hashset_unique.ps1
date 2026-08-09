# vybe-test: powershell/generic_types/generic_hashset_unique
$set = [System.Collections.Generic.HashSet[int]]::new()
[void]$set.Add(1)
[void]$set.Add(1)
[void]$set.Add(2)
if ($set.Count -ne 2) {
    Write-Host "FAIL: HashSet[int] Count expected 2 unique items, got $($set.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
