# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_create_three_items
$t = [System.Tuple]::Create(1, "two", 3.0)
if ($t.Item1 -ne 1 -or $t.Item2 -ne "two" -or $t.Item3 -ne 3.0) {
    Write-Host "FAIL: 3-tuple creation failed"
    exit 1
}
Write-Host "PASS"
exit 0
