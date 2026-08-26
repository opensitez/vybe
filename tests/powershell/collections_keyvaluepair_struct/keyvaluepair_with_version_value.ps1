# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_with_version_value
$v = [version]"7.4.0"
$kvp = [System.Collections.Generic.KeyValuePair[string, version]]::new("version", $v)
if ($kvp.Key -ne "version" -or $kvp.Value.Major -ne 7) {
    Write-Host "FAIL: Version value KeyValuePair failed"
    exit 1
}
Write-Host "PASS"
exit 0
