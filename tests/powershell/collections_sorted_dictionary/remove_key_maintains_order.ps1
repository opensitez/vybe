# vybe-test: powershell/collections_sorted_dictionary/remove_key_maintains_order
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("a", 1); $sd.Add("b", 2); $sd.Add("c", 3)
$rem = $sd.Remove("b")
$keys = @($sd.Keys)
if (-not $rem -or $keys.Count -ne 2 -or $keys[0] -ne "a" -or $keys[1] -ne "c") {
    Write-Host "FAIL: Remove key maintaining order failed"
    exit 1
}
Write-Host "PASS"
exit 0
