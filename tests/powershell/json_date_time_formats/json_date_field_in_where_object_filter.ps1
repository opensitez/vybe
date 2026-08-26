# vybe-test: powershell/json_date_time_formats/json_date_field_in_where_object_filter
$events = @(
    '{"id":1,"date":"2026-01-10"}',
    '{"id":2,"date":"2026-09-15"}'
) | ConvertFrom-Json
$target = [datetime]::Parse("2026-06-01")
$filtered = @($events | Where-Object { [datetime]::Parse($_.date) -gt $target })
if ($filtered.Length -ne 1 -or $filtered[0].id -ne 2) {
    Write-Host "FAIL: JSON date field in Where-Object filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
