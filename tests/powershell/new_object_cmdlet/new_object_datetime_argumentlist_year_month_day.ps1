# vybe-test: powershell/new_object_cmdlet/new_object_datetime_argumentlist_year_month_day
# New-Object with -ArgumentList @(year, month, day) constructs a DateTime with correct components
$dt = New-Object System.DateTime -ArgumentList 2030, 6, 15

if ($dt.Year -ne 2030 -or $dt.Month -ne 6 -or $dt.Day -ne 15) {
    Write-Host "FAIL: DateTime components mismatch: $($dt.ToString('yyyy-MM-dd'))"
    exit 1
}

Write-Host "PASS"
exit 0
