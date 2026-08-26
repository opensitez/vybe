# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_static_property_lookup
$prop = "PI"
$res = [math]::$prop
if ([math]::Abs($res - 3.14159265358979) -gt 1e-10) {
    Write-Host "FAIL: Dynamic static property lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
