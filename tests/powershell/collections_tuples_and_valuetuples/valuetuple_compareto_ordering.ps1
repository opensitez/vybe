# vybe-test: powershell/collections_tuples_and_valuetuples/valuetuple_compareto_ordering
$vt1 = [System.ValueTuple]::Create(1, "apple")
$vt2 = [System.ValueTuple]::Create(1, "banana")
if ($vt1.CompareTo($vt2) -ge 0) {
    Write-Host "FAIL: ValueTuple CompareTo ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
