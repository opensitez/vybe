# vybe-test: powershell/collections_readonly_collections/readonly_dictionary_with_guid_keys
$d = [System.Collections.Generic.Dictionary[guid, string]]::new()
$g = [guid]::NewGuid()
$d.Add($g, "GuidData")
$rod = [System.Collections.ObjectModel.ReadOnlyDictionary[guid, string]]::new($d)
if ($rod[$g] -ne "GuidData") { Write-Host "FAIL: Guid dictionary failed"; exit 1 }
Write-Host "PASS"; exit 0
