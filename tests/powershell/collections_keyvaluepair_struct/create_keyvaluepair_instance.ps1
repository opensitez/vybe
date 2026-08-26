# vybe-test: powershell/collections_keyvaluepair_struct/create_keyvaluepair_instance
$kvp = [System.Collections.Generic.KeyValuePair[string, int]]::new("score", 100)
if ($kvp.Key -ne "score" -or $kvp.Value -ne 100) {
    Write-Host "FAIL: KeyValuePair construction failed"
    exit 1
}
Write-Host "PASS"
exit 0
