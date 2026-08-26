# vybe-test: powershell/classes_custom_indexer_properties/class_indexer_access_14
class IndexableStore_14 {
    [hashtable]$Data = @{}
    [string]GetItem([string]$key) { return $this.Data[$key] }
    [void]SetItem([string]$key, [string]$val) { $this.Data[$key] = $val }
}
$store = [IndexableStore_14]::new()
$store.SetItem("name", "Vybe_14")
if ($store.GetItem("name") -ne "Vybe_14") { Write-Host "FAIL: Custom indexer failed"; exit 1 }
Write-Host "PASS"; exit 0
