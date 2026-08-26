# vybe-test: powershell/collections_tuples_and_valuetuples/valuetuple_tostring_format
$vt = [System.ValueTuple]::Create("a", "b")
if ($vt.ToString() -ne "(a, b)") {
    Write-Host "FAIL: ValueTuple ToString failed, expected '(a, b)', got '$($vt.ToString())'"
    exit 1
}
Write-Host "PASS"
exit 0
