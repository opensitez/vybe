# vybe-test: powershell/collections_generic_hashset/foreach_enumeration
$set = [System.Collections.Generic.HashSet[int]]::new([int[]]@(10, 20, 30))
$sum = 0
foreach ($item in $set) { $sum += $item }
if ($sum -ne 60) {
    Write-Host "FAIL: foreach on HashSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
