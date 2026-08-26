# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_hashcode_consistency
$t1 = [System.Tuple]::Create("a", "b")
$t2 = [System.Tuple]::Create("a", "b")
if ($t1.GetHashCode() -ne $t2.GetHashCode()) {
    Write-Host "FAIL: Tuple HashCode consistency failed"
    exit 1
}
Write-Host "PASS"
exit 0
