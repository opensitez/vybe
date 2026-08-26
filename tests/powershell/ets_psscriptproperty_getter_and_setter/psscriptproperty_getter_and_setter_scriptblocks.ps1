# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_getter_and_setter_scriptblocks
$obj = [pscustomobject]@{ Celsius = 0.0 }
$prop = [System.Management.Automation.PSScriptProperty]::new(
    "Fahrenheit",
    { ($this.Celsius * 9.0 / 5.0) + 32.0 },
    { param($f) $this.Celsius = ($f - 32.0) * 5.0 / 9.0 }
)
$obj.PSObject.Properties.Add($prop)
$f0 = $obj.Fahrenheit # 32
$obj.Fahrenheit = 212.0
if ($f0 -ne 32.0 -or $obj.Celsius -ne 100.0 -or $obj.Fahrenheit -ne 212.0) {
    Write-Host "FAIL: PSScriptProperty getter and setter failed, f0=$f0, c=$($obj.Celsius)"
    exit 1
}
Write-Host "PASS"
exit 0
