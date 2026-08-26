# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_in_generic_list
$list = [System.Collections.Generic.List[System.Tuple[string, int]]]::new()
$list.Add([System.Tuple]::Create("one", 1))
$list.Add([System.Tuple]::Create("two", 2))
if ($list.Count -ne 2 -or $list[1].Item1 -ne "two") {
    Write-Host "FAIL: List of Tuples failed"
    exit 1
}
Write-Host "PASS"
exit 0
