# vybe-test: powershell/loops/foreach_hashtable_pairs
$scores = @{ Alice = 90; Bob = 85; Carol = 92 }
$total = 0
$count = 0
foreach ($entry in $scores.GetEnumerator()) {
    $total += $entry.Value
    $count++
}
if ($count -ne 3)   { Write-Host "FAIL: count $count"; exit 1 }
if ($total -ne 267) { Write-Host "FAIL: total $total"; exit 1 }
Write-Host "PASS"
exit 0
