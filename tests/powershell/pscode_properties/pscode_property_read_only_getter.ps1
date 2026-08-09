# vybe-test: powershell/pscode_properties/pscode_property_read_only_getter
class ReadOnlyCodeHelper {
    static [string] GetConstant([object]$t) { return "Immutable" }
}
$obj = [pscustomobject]@{}
$g = [ReadOnlyCodeHelper].GetMethod("GetConstant")
$cp = [System.Management.Automation.PSCodeProperty]::new("ConstProp", $g)
$obj.psobject.Members.Add($cp)
if (-not $cp.IsGettable -or $cp.IsSettable) {
    Write-Host "FAIL: read-only PSCodeProperty expected IsGettable=true, IsSettable=false"
    exit 1
}
Write-Host "PASS"
exit 0
