# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_iso8601_roundtrip_format
# Get-Date -Format 'o' generates an ISO 8601 round-trip date-time string
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$iso = Get-Date -Date $date -Format "o"

if ($iso -notmatch "^2026-05-10T14:30:45") {
    Write-Host "FAIL: ISO 8601 pattern mismatch: '$iso'"
    exit 1
}

Write-Host "PASS"
exit 0
