# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_read_on_guid
$prop = "Guid"
$g = [guid]::NewGuid()
$res = $g.$prop
if ($res -ne $g.ToString()) {
    Write-Host "FAIL: Dynamic property read on GUID failed"
    exit 1
}
Write-Host "PASS"
exit 0
