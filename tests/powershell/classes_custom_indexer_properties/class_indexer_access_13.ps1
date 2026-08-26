# vybe-test: powershell/classes_custom_indexer_properties/class_indexer_access_13
class IndexableStore_13 {
    [hashtable]$Data = @{}
    [string]GetItem([string]$key) { return $this.Data[$key] }
    [void]SetItem([string]$key, [string]$val) { $this.Data[$key] = $val }
}
$store = [IndexableStore_13]::new()
$store.SetItem("name", "Vybe_13")
if ($store.GetItem("name") -ne "Vybe_13") { Write-Host "FAIL: Custom indexer failed"; exit 1 }
Write-Host "PASS"; exit 0
