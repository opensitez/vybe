# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_with_object_value
$obj = [pscustomobject]@{ Id = 10 }
$kvp = [System.Collections.Generic.KeyValuePair[string, object]]::new("ObjKey", $obj)
if ($kvp.Key -ne "ObjKey" -or $kvp.Value.Id -ne 10) {
    Write-Host "FAIL: KeyValuePair with object value failed"
    exit 1
}
Write-Host "PASS"
exit 0
