# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_equality_same_values
$t1 = [System.Tuple]::Create("key", 100)
$t2 = [System.Tuple]::Create("key", 100)
if ($t1 -ne $t2) {
    Write-Host "FAIL: Identical tuples must be equal"
    exit 1
}
Write-Host "PASS"
exit 0
