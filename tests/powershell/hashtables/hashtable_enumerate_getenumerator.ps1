# vybe-test: powershell/hashtables/hashtable_enumerate_getEnumerator
$h = @{ a = 1; b = 2; c = 3 }
$sum = 0
foreach ($pair in $h.GetEnumerator()) {
    $sum += $pair.Value
}
if ($sum -ne 6) {
    Write-Host "FAIL: expected sum 6, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
