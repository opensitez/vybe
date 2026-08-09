# vybe-test: powershell/type_converters/type_converter_string_to_datetime
$dt = [datetime]"2026-01-01"
if ($dt.Year -ne 2026 -or $dt.Month -ne 1 -or $dt.Day -ne 1) {
    Write-Host "FAIL: datetime converter expected 2026-01-01, got $($dt.ToString('yyyy-MM-dd'))"
    exit 1
}
Write-Host "PASS"
exit 0
