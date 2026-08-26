# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_int_key
$kvp = [System.Collections.Generic.KeyValuePair[int, string]]::new(404, "NotFound")
if ($kvp.Key -ne 404 -or $kvp.Value -ne "NotFound") {
    Write-Host "FAIL: Integer key KeyValuePair failed"
    exit 1
}
Write-Host "PASS"
exit 0
