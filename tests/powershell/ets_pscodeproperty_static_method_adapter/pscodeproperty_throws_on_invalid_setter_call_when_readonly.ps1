# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_throws_on_invalid_setter_call_when_readonly
class RoClass {
    static [int]GetNum([psobject]$i) { return 10 }
}
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("NumRO", [RoClass].GetMethod("GetNum")))
$caught = $false
try {
    $obj.NumRO = 99
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Assigning to read-only PSCodeProperty must throw"
    exit 1
}
Write-Host "PASS"
exit 0
