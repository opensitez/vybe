# vybe-test: powershell/collections_tuples_and_valuetuples/valuetuple_create_two_items
$vt = [System.ValueTuple]::Create("Bob", 25)
if ($vt.Item1 -ne "Bob" -or $vt.Item2 -ne 25) {
    Write-Host "FAIL: ValueTuple create failed"
    exit 1
}
Write-Host "PASS"
exit 0
