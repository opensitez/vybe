# vybe-test: powershell/collections_tuples_and_valuetuples/valuetuple_single_element
$vt = [System.ValueTuple]::Create(42)
if ($vt.Item1 -ne 42) {
    Write-Host "FAIL: 1-item ValueTuple failed"
    exit 1
}
Write-Host "PASS"
exit 0
