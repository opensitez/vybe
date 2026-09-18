# vybe-test: powershell/out_string_cmdlet/out_string_datetime_formatting
# DateTime instances stringify through Out-String adhering to system culture date/time formatting
$dt = [DateTime]::Parse("2026-07-04 15:30:00")
$output = $dt | Out-String

if ($output -notmatch "2026" -or $output -notmatch "July") {
    Write-Host "FAIL: date formatting mismatch in Out-String output: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
