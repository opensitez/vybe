# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_on_datetime_instance
$dt = [datetime]::Parse("2026-08-26")
$m = "AddDays"
$res = $dt.$m(5)
if ($res.Day -ne 31) {
    Write-Host "FAIL: Dynamic method on DateTime failed"
    exit 1
}
Write-Host "PASS"
exit 0
