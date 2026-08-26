# vybe-test: powershell/type_hashcode_combiner/hashcode_with_custom_class_instance
class HashKeyItem {
    [int]$Id
    [string]$Name
    HashKeyItem([int]$i, [string]$n) { $this.Id = $i; $this.Name = $n }
    [int]GetHashCode() { return [System.HashCode]::Combine($this.Id, $this.Name) }
}
$k1 = [HashKeyItem]::new(1, "Item")
$k2 = [HashKeyItem]::new(1, "Item")
if ($k1.GetHashCode() -ne $k2.GetHashCode()) { Write-Host "FAIL: Custom class HashCode combine failed"; exit 1 }
Write-Host "PASS"; exit 0
