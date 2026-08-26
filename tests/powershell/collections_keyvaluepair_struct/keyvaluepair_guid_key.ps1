# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_guid_key
$g = [guid]::NewGuid()
$kvp = [System.Collections.Generic.KeyValuePair[guid, bool]]::new($g, $true)
if ($kvp.Key -ne $g -or $kvp.Value -ne $true) {
    Write-Host "FAIL: Guid key KeyValuePair failed"
    exit 1
}
Write-Host "PASS"
exit 0
