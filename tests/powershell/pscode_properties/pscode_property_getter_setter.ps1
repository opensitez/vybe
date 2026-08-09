# vybe-test: powershell/pscode_properties/pscode_property_getter_setter
class StorageHelper {
    static [string]$Backing = "Init"
    static [string] GetStorage([object]$t) { return [StorageHelper]::$Backing }
    static [void] SetStorage([object]$t, [string]$val) { [StorageHelper]::$Backing = $val }
}
$obj = [pscustomobject]@{}
$g = [StorageHelper].GetMethod("GetStorage")
$s = [StorageHelper].GetMethod("SetStorage")
$cp = [System.Management.Automation.PSCodeProperty]::new("Prop", $g, $s)
$obj.psobject.Members.Add($cp)
$obj.Prop = "Mutated"
if ($obj.Prop -ne "Mutated") {
    Write-Host "FAIL: PSCodeProperty getter/setter expected Mutated, got $($obj.Prop)"
    exit 1
}
Write-Host "PASS"
exit 0
