# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_inequality_different_values
$t1 = [System.Tuple]::Create(1, 2)
$t2 = [System.Tuple]::Create(1, 3)
if ($t1 -eq $t2) {
    Write-Host "FAIL: Different tuples must compare unequal"
    exit 1
}
Write-Host "PASS"
exit 0
