# vybe-test: powershell/classes_constructor_overloading/constructor_hashtable_parameter
class Bag {
    [hashtable]$Data
    Bag([hashtable]$ht) { $this.Data = $ht }
}
$b = [Bag]::new(@{ key = "value" })
if ($b.Data["key"] -ne "value") {
    Write-Host "FAIL: Hashtable parameter constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
