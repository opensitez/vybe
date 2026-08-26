# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_returns_array
$obj = [pscustomobject]@{ Csv = "10,20,30" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Numbers", { @($this.Csv.Split(',') | ForEach-Object { [int]$_ }) }))
if ($obj.Numbers.Length -ne 3 -or $obj.Numbers[0] -ne 10 -or $obj.Numbers[2] -ne 30) {
    Write-Host "FAIL: PSScriptProperty returning array failed"
    exit 1
}
Write-Host "PASS"
exit 0
