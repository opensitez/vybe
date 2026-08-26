# vybe-test: powershell/collections_keyvaluepair_struct/dictionary_enumeration_yields_keyvaluepair
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("testKey", 42)
$keys = @($d.Keys)
if ($keys[0] -ne "testKey" -or $d["testKey"] -ne 42) {
    Write-Host "FAIL: Dictionary Key/Value access failed"
    exit 1
}
Write-Host "PASS"
exit 0
