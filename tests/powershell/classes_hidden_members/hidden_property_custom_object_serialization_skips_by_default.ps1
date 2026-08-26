# vybe-test: powershell/classes_hidden_members/hidden_property_custom_object_serialization_skips_by_default
class ExportItem {
    [string]$Public = "visible"
    hidden [string]$Private = "hidden"
}
$item = [ExportItem]::new()
$props = @($item.PSObject.Properties | ForEach-Object { $_.Name })
if ($props -contains "Private" -or -not ($props -contains "Public")) {
    Write-Host "FAIL: Hidden property should be excluded from default PSObject.Properties"
    exit 1
}
Write-Host "PASS"
exit 0
