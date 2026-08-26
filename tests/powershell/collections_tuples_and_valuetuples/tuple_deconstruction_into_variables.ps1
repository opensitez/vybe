# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_deconstruction_into_variables
$t = [System.Tuple]::Create("first", 999)
$val1 = $t.Item1
$val2 = $t.Item2
if ($val1 -ne "first" -or $val2 -ne 999) {
    Write-Host "FAIL: Manual deconstruction failed"
    exit 1
}
Write-Host "PASS"
exit 0
