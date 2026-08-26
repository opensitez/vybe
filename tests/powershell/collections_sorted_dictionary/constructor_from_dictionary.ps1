# vybe-test: powershell/collections_sorted_dictionary/constructor_from_dictionary
$orig = [System.Collections.Generic.Dictionary[string, int]]::new()
$orig.Add("c", 3); $orig.Add("a", 1); $orig.Add("b", 2)
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new($orig)
$keys = @($sd.Keys)
if ($keys[0] -ne "a" -or $keys[2] -ne "c") {
    Write-Host "FAIL: Constructor from IDictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
