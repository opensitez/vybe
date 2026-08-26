# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_in_generic_list
$list = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string, int]]]::new()
$list.Add([System.Collections.Generic.KeyValuePair[string, int]]::new("a", 1))
$list.Add([System.Collections.Generic.KeyValuePair[string, int]]::new("b", 2))
if ($list.Count -ne 2 -or $list[1].Key -ne "b") {
    Write-Host "FAIL: List of KeyValuePairs failed"
    exit 1
}
Write-Host "PASS"
exit 0
