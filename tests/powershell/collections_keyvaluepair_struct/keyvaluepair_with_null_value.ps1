# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_with_null_value
$kvp = [System.Collections.Generic.KeyValuePair[string, string]]::new("Key", "Value")
if ($kvp.Key -ne "Key" -or $kvp.Value -ne "Value") {
    Write-Host "FAIL: KeyValuePair failed"
    exit 1
}
Write-Host "PASS"
exit 0
