# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_with_boolean_value
$kvp = [System.Collections.Generic.KeyValuePair[string, bool]]::new("isEnabled", $true)
if ($kvp.Key -ne "isEnabled" -or $kvp.Value -ne $true) {
    Write-Host "FAIL: Boolean value KeyValuePair failed"
    exit 1
}
Write-Host "PASS"
exit 0
