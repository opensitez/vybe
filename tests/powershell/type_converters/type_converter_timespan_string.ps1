# vybe-test: powershell/type_converters/type_converter_timespan_string
$ts = [timespan]"01:00:00"
if ($ts.TotalSeconds -ne 3600) {
    Write-Host "FAIL: string to [timespan] conversion expected TotalSeconds=3600, got $($ts.TotalSeconds)"
    exit 1
}
Write-Host "PASS"
exit 0
