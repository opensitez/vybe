# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_create_two_items
$t = [System.Tuple]::Create("Alice", 30)
if ($t.Item1 -ne "Alice" -or $t.Item2 -ne 30) {
    Write-Host "FAIL: 2-tuple creation failed"
    exit 1
}
Write-Host "PASS"
exit 0
