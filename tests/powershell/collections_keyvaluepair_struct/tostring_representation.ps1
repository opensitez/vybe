# vybe-test: powershell/collections_keyvaluepair_struct/tostring_representation
$kvp = [System.Collections.Generic.KeyValuePair[string, string]]::new("env", "production")
if ($kvp.ToString() -ne "[env, production]") {
    Write-Host "FAIL: KeyValuePair ToString failed, expected '[env, production]', got '$($kvp.ToString())'"
    exit 1
}
Write-Host "PASS"
exit 0
