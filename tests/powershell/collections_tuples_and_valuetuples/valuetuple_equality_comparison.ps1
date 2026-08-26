# vybe-test: powershell/collections_tuples_and_valuetuples/valuetuple_equality_comparison
$vt1 = [System.ValueTuple]::Create(10, 20)
$vt2 = [System.ValueTuple]::Create(10, 20)
if ($vt1 -ne $vt2) {
    Write-Host "FAIL: ValueTuples with same items must be equal"
    exit 1
}
Write-Host "PASS"
exit 0
