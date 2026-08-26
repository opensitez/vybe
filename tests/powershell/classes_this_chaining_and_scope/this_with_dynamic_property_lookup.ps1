# vybe-test: powershell/classes_this_chaining_and_scope/this_with_dynamic_property_lookup
class DynamicPropClass {
    [string]$Title = "Main"
    [string]GetField([string]$propName) {
        return $this.$propName
    }
}
$dpc = [DynamicPropClass]::new()
$val = $dpc.GetField("Title")
if ($val -ne "Main") {
    Write-Host "FAIL: Dynamic property access via `$this failed, got '$val'"
    exit 1
}
Write-Host "PASS"
exit 0
