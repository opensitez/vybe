# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_immutability
$kvp = [System.Collections.Generic.KeyValuePair[string, int]]::new("fixed", 50)
if ($kvp.Key -ne "fixed" -or $kvp.Value -ne 50) {
    Write-Host "FAIL: Initial values check failed"
    exit 1
}
Write-Host "PASS"
exit 0
