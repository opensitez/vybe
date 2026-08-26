# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_read_on_dotnet_type
$prop = "Year"
$dt = [datetime]::Parse("2026-08-26")
$res = $dt.$prop
if ($res -ne 2026) {
    Write-Host "FAIL: Dynamic property read on .NET instance failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
