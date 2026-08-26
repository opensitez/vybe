# vybe-test: powershell/json_enum_and_primitive_serialization/enum_in_custom_class_property_serialization
enum StatusType { Pending; Active; Closed }
class ItemStatus {
    [StatusType]$Status = [StatusType]::Active
}
$is = [ItemStatus]::new()
$json = $is | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Status -ne 1 -and $recovered.Status -ne "Active") {
    Write-Host "FAIL: Enum in custom class property serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
