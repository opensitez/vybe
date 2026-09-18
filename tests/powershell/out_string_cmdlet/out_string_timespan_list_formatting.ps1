# vybe-test: powershell/out_string_cmdlet/out_string_timespan_list_formatting
# TimeSpan instances format through Out-String in property-list layout
$ts = [TimeSpan]::FromMinutes(90)
$output = $ts | Out-String

if ($output -notmatch "TotalMinutes\s*:\s*90") {
    Write-Host "FAIL: TotalMinutes missing from TimeSpan formatting: '$output'"
    exit 1
}

if ($output -notmatch "Hours\s*:\s*1") {
    Write-Host "FAIL: Hours missing from TimeSpan formatting: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
