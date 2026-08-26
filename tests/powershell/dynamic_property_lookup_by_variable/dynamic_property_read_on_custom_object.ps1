# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_read_on_custom_object
$prop = "ServerName"
$obj = [pscustomobject]@{ ServerName = "db-prod-01"; Port = 5432 }
$res = $obj.$prop
if ($res -ne "db-prod-01") {
    Write-Host "FAIL: Dynamic property read on custom object failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
