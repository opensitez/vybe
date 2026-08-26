# vybe-test: powershell/json_date_time_formats/leap_year_date_in_json
$json = '{"LeapDay":"2028-02-29"}'
$obj = $json | ConvertFrom-Json
$dt = [datetime]::ParseExact($obj.LeapDay, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
if ($dt.Month -ne 2 -or $dt.Day -ne 29) {
    Write-Host "FAIL: Leap year date in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
