# vybe-test: powershell/collections_sorted_dictionary/containskey_check
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("target", 1)
if (-not $sd.ContainsKey("target") -or $sd.ContainsKey("missing")) {
    Write-Host "FAIL: ContainsKey check failed"
    exit 1
}
Write-Host "PASS"
exit 0
