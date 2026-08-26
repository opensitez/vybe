# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_conversion_from_hashtable_entry
$ht = @{ status = "ok" }
$entry = $ht.GetEnumerator() | Select-Object -First 1
$kvp = [System.Collections.Generic.KeyValuePair[string, string]]::new([string]$entry.Key, [string]$entry.Value)
if ($kvp.Key -ne "status" -or $kvp.Value -ne "ok") {
    Write-Host "FAIL: Conversion from DictionaryEntry to KeyValuePair failed"
    exit 1
}
Write-Host "PASS"
exit 0
