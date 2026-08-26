# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_read_on_version
$prop = "Major"
$v = [version]"7.4.1"
$res = $v.$prop
if ($res -ne 7) {
    Write-Host "FAIL: Dynamic property read on Version failed"
    exit 1
}
Write-Host "PASS"
exit 0
