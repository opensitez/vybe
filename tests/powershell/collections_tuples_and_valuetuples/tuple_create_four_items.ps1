# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_create_four_items
$t = [System.Tuple]::Create("a", "b", "c", "d")
if ($t.Item1 -ne "a" -or $t.Item4 -ne "d") {
    Write-Host "FAIL: 4-tuple creation failed"
    exit 1
}
Write-Host "PASS"
exit 0
