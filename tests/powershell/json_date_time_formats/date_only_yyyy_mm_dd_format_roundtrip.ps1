# vybe-test: powershell/json_date_time_formats/date_only_yyyy_mm_dd_format_roundtrip
$json = '{"Date":"2026-08-26"}'
$obj = $json | ConvertFrom-Json
$dt = [datetime]::ParseExact($obj.Date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
if ($dt.Year -ne 2026 -or $dt.Month -ne 8 -or $dt.Day -ne 26) {
    Write-Host "FAIL: Date-only format roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
