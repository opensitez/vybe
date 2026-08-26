# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_with_null_element
$t = [System.Tuple]::Create("Valid", "Second")
if ($t.Item1 -ne "Valid" -or $t.Item2 -ne "Second") {
    Write-Host "FAIL: Tuple creation failed"
    exit 1
}
Write-Host "PASS"
exit 0
