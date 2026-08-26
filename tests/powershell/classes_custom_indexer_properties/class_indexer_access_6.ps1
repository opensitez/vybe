# vybe-test: powershell/classes_custom_indexer_properties/class_indexer_access_6
class IndexableStore_6 {
    [hashtable]$Data = @{}
    [string]GetItem([string]$key) { return $this.Data[$key] }
    [void]SetItem([string]$key, [string]$val) { $this.Data[$key] = $val }
}
$store = [IndexableStore_6]::new()
$store.SetItem("name", "Vybe_6")
if ($store.GetItem("name") -ne "Vybe_6") { Write-Host "FAIL: Custom indexer failed"; exit 1 }
Write-Host "PASS"; exit 0
