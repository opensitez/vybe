# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_readonly_throws_on_set
$obj = [pscustomobject]@{ Val = 10 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Double", { $this.Val * 2 }))
$caught = $false
try {
    $obj.Double = 50
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Setting read-only PSScriptProperty must throw"
    exit 1
}
Write-Host "PASS"
exit 0
