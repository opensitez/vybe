# vybe-test: powershell/collections_tuples_and_valuetuples/nested_tuple_structure
$inner = [System.Tuple]::Create(1, 2)
$outer = [System.Tuple]::Create("point", $inner)
if ($outer.Item1 -ne "point" -or $outer.Item2.Item1 -ne 1 -or $outer.Item2.Item2 -ne 2) {
    Write-Host "FAIL: Nested Tuple structure failed"
    exit 1
}
Write-Host "PASS"
exit 0
