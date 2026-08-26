# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_write_on_custom_object
$prop = "Status"
$obj = [pscustomobject]@{ Status = "Offline" }
$obj.$prop = "Online"
if ($obj.Status -ne "Online" -or $obj.$prop -ne "Online") {
    Write-Host "FAIL: Dynamic property write failed"
    exit 1
}
Write-Host "PASS"
exit 0
