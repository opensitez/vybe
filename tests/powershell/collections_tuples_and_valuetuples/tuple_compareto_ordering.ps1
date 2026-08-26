# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_compareto_ordering
$t1 = [System.Tuple]::Create(1, 10)
$t2 = [System.Tuple]::Create(1, 20)
if ($t1.CompareTo($t2) -ge 0) {
    Write-Host "FAIL: Tuple CompareTo ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
