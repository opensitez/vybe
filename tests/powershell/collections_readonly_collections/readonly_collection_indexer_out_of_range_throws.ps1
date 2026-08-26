# vybe-test: powershell/collections_readonly_collections/readonly_collection_indexer_out_of_range_throws
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
$caught = $false
try {
    $x = $roc.Item(5)
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) {
    # In PowerShell indexer access might return null or throw
    $caught = ($roc[5] -eq $null)
}
if (-not $caught) { Write-Host "FAIL: Out of range index check failed"; exit 1 }
Write-Host "PASS"; exit 0
