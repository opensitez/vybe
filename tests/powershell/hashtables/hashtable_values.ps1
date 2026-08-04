# vybe-test: powershell/hashtables/hashtable_values
$hash = @{ A = 10; B = 20; C = 30 }
$values = $hash.Values
$sum = 0
foreach ($val in $values) {
    $sum += $val
}
if ($sum -ne 60) {
    Write-Host "FAIL: expected 60, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
