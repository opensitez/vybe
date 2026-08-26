# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_create_seven_items
$t = [System.Tuple]::Create(1, 2, 3, 4, 5, 6, 7)
if ($t.Item1 -ne 1 -or $t.Item7 -ne 7) {
    Write-Host "FAIL: 7-tuple creation failed"
    exit 1
}
Write-Host "PASS"
exit 0
