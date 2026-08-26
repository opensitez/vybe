# vybe-test: powershell/collections_sorted_dictionary/copyto_keyvaluepair_array
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("k", 10)
$arr = [System.Array]::CreateInstance([type]"System.Collections.Generic.KeyValuePair[string, int]", 1)
$sd.CopyTo($arr, 0)
if ($arr[0].Key -ne "k" -or $arr[0].Value -ne 10) {
    Write-Host "FAIL: CopyTo KeyValuePair array failed"
    exit 1
}
Write-Host "PASS"
exit 0
