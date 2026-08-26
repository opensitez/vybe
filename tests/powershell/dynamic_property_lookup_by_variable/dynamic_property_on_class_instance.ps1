# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_on_class_instance
class TargetEntity {
    [string]$Label = "MyEntity"
    [int]$Rank = 1
}
$te = [TargetEntity]::new()
$p1 = "Label"
$p2 = "Rank"
if ($te.$p1 -ne "MyEntity" -or $te.$p2 -ne 1) {
    Write-Host "FAIL: Dynamic property on class instance failed"
    exit 1
}
Write-Host "PASS"
exit 0
