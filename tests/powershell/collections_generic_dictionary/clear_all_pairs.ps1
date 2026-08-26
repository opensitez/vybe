# vybe-test: powershell/collections_generic_dictionary/clear_all_pairs
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("k", 1)
$d.Clear()
if ($d.Count -ne 0) {
    Write-Host "FAIL: Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
