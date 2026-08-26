# vybe-test: powershell/collections_generic_dictionary/add_key_value_pairs
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("one", 1)
$d.Add("two", 2)
if ($d.Count -ne 2 -or $d["one"] -ne 1 -or $d["two"] -ne 2) {
    Write-Host "FAIL: Dictionary Add failed"
    exit 1
}
Write-Host "PASS"
exit 0
