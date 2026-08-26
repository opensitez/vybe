# vybe-test: powershell/collections_readonly_collections/readonly_collection_contains_and_indexof
$list = [System.Collections.Generic.List[string]]::new([string[]]@("a", "b", "c"))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[string]]::new($list)
if (-not $roc.Contains("b") -or $roc.IndexOf("c") -ne 2) { Write-Host "FAIL: Contains/IndexOf failed"; exit 1 }
Write-Host "PASS"; exit 0
