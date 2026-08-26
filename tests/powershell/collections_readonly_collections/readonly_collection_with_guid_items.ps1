# vybe-test: powershell/collections_readonly_collections/readonly_collection_with_guid_items
$list = [System.Collections.Generic.List[guid]]::new()
$g = [guid]::NewGuid()
$list.Add($g)
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[guid]]::new($list)
if ($roc[0] -ne $g) { Write-Host "FAIL: Guid collection failed"; exit 1 }
Write-Host "PASS"; exit 0
