# vybe-test: powershell/json_date_time_formats/timespan_string_format_c_in_json
$ts = [timespan]::FromSeconds(3723) # 01:02:03
$json = @{ Duration = $ts.ToString("c") } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
$recovered = [timespan]::Parse($obj.Duration)
if ($recovered.TotalSeconds -ne 3723) {
    Write-Host "FAIL: TimeSpan 'c' string format in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
