# vybe-test: powershell/json_date_time_formats/date_comparison_after_json_deserialization
$json = '{"d1":"2026-01-01","d2":"2026-06-01"}'
$obj = $json | ConvertFrom-Json
$t1 = [datetime]::ParseExact($obj.d1, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
$t2 = [datetime]::ParseExact($obj.d2, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
if (-not ($t1 -lt $t2)) {
    Write-Host "FAIL: Date comparison after deserialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
