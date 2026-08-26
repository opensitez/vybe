# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_tostring_format
$t = [System.Tuple]::Create(1, 2)
if ($t.ToString() -ne "(1, 2)") {
    Write-Host "FAIL: Tuple ToString failed, expected '(1, 2)', got '$($t.ToString())'"
    exit 1
}
Write-Host "PASS"
exit 0
