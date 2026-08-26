# vybe-test: powershell/collections_generic_dictionary/values_collection_enumeration
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("k1", 10); $d.Add("k2", 20)
$sum = 0
foreach ($v in $d.Values) { $sum += $v }
if ($sum -ne 30) {
    Write-Host "FAIL: Values collection enumeration failed"
    exit 1
}
Write-Host "PASS"
exit 0
