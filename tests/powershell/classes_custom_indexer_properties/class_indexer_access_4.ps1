# vybe-test: powershell/classes_custom_indexer_properties/class_indexer_access_4
class IndexableStore_4 {
    [hashtable]$Data = @{}
    [string]GetItem([string]$key) { return $this.Data[$key] }
    [void]SetItem([string]$key, [string]$val) { $this.Data[$key] = $val }
}
$store = [IndexableStore_4]::new()
$store.SetItem("name", "Vybe_4")
if ($store.GetItem("name") -ne "Vybe_4") { Write-Host "FAIL: Custom indexer failed"; exit 1 }
Write-Host "PASS"; exit 0
