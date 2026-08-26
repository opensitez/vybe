# vybe-test: powershell/collections_tuples_and_valuetuples/valuetuple_three_elements
$vt = [System.ValueTuple]::Create(1.5, "test", $true)
if ($vt.Item1 -ne 1.5 -or $vt.Item2 -ne "test" -or $vt.Item3 -ne $true) {
    Write-Host "FAIL: ValueTuple 3 elements failed"
    exit 1
}
Write-Host "PASS"
exit 0
