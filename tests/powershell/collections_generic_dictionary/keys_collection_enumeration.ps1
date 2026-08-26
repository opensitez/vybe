# vybe-test: powershell/collections_generic_dictionary/keys_collection_enumeration
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("k1", 1); $d.Add("k2", 2)
$keys = @($d.Keys)
if ($keys.Count -ne 2 -or -not ($keys -contains "k1") -or -not ($keys -contains "k2")) {
    Write-Host "FAIL: Keys collection extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
