# vybe-test: powershell/collections_sorted_dictionary/numeric_key_sorting
$sd = [System.Collections.Generic.SortedDictionary[int, string]]::new()
$sd.Add(50, "fifty"); $sd.Add(10, "ten"); $sd.Add(30, "thirty")
$keys = @($sd.Keys)
if ($keys[0] -ne 10 -or $keys[1] -ne 30 -or $keys[2] -ne 50) {
    Write-Host "FAIL: Integer key sorting failed"
    exit 1
}
Write-Host "PASS"
exit 0
