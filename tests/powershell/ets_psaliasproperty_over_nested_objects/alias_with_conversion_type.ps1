# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_with_conversion_type
$obj = [pscustomobject]@{ PortStr = "8080" }
$alias = [System.Management.Automation.PSAliasProperty]::new("PortInt", "PortStr", [int])
$obj.PSObject.Properties.Add($alias)
if ($obj.PortInt -ne 8080 -or $obj.PortInt -isnot [int]) {
    Write-Host "FAIL: PSAliasProperty with type conversion failed"
    exit 1
}
Write-Host "PASS"
exit 0
