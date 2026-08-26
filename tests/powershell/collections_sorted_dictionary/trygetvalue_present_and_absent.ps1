# vybe-test: powershell/collections_sorted_dictionary/trygetvalue_present_and_absent
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("val", 777)
$hasKey = $sd.ContainsKey("val")
$val = if ($hasKey) { $sd["val"] } else { 0 }
if (-not $hasKey -or $val -ne 777 -or $sd.ContainsKey("none")) {
    Write-Host "FAIL: SortedDictionary lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
