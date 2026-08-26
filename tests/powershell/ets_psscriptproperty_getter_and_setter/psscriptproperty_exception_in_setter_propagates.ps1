# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_exception_in_setter_propagates
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("StrictVal", { 1 }, { param($v) if ($v -lt 0) { throw "NegativeNotAllowed" } }))
$caught = $false
try {
    $obj.StrictVal = -10
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: PSScriptProperty setter exception propagation failed"
    exit 1
}
Write-Host "PASS"
exit 0
