# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_static_guid_empty_lookup
$prop = "Empty"
$res = [guid]::$prop
if ($res -ne [guid]::Empty) {
    Write-Host "FAIL: Dynamic static GUID.Empty lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
