# vybe-test: powershell/collections_generic_dictionary/indexer_get_and_set
$d = [System.Collections.Generic.Dictionary[string, string]]::new()
$d["key1"] = "val1"
$d["key1"] = "updated"
if ($d["key1"] -ne "updated" -or $d.Count -ne 1) {
    Write-Host "FAIL: Dictionary indexer update failed"
    exit 1
}
Write-Host "PASS"
exit 0
