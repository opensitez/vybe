# vybe-test: powershell/pscustomobject_literals/pscustomobject_property_names
$obj = [pscustomobject]@{ Alpha = 1; Beta = 2 }
$props = @($obj.psobject.Properties.Name)
if ($props[0] -ne "Alpha" -or $props[1] -ne "Beta") {
    Write-Host "FAIL: property order expected Alpha, Beta, got $($props -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
