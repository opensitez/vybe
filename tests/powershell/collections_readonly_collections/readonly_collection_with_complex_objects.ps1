# vybe-test: powershell/collections_readonly_collections/readonly_collection_with_complex_objects
$list = [System.Collections.Generic.List[pscustomobject]]::new()
$list.Add([pscustomobject]@{ Name = "ReadOnlyObj" })
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[pscustomobject]]::new($list)
if ($roc[0].Name -ne "ReadOnlyObj") { Write-Host "FAIL: Complex object failed"; exit 1 }
Write-Host "PASS"; exit 0
