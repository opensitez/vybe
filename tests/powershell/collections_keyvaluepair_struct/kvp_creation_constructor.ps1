# vybe-test: powershell/collections_keyvaluepair_struct/kvp_creation_constructor
$kvp = [System.Collections.Generic.KeyValuePair[string, int]]::new("apples", 5)
if ($kvp.Key -ne "apples" -or $kvp.Value -ne 5) {
    Write-Host "FAIL: KeyValuePair constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
