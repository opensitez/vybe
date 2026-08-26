# vybe-test: powershell/collections_generic_hashset/clear_all
$set = [System.Collections.Generic.HashSet[string]]::new([string[]]@("a", "b"))
$set.Clear()
if ($set.Count -ne 0) {
    Write-Host "FAIL: HashSet Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
