# vybe-test: powershell/collections_generic_hashset/trimexcess_capacity
$set = [System.Collections.Generic.HashSet[int]]::new(100)
$set.Add(1); $set.Add(2)
$set.TrimExcess()
if ($set.Count -ne 2) {
    Write-Host "FAIL: TrimExcess failed"
    exit 1
}
Write-Host "PASS"
exit 0
