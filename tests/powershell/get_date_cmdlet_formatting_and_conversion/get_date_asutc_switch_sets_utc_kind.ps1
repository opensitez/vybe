# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_asutc_switch_sets_utc_kind
# The -AsUTC switch converts the output datetime into Universal Coordinated Time with DateTimeKind.Utc
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$utc = Get-Date -Date $date -AsUTC

if ($utc.Kind -ne [System.DateTimeKind]::Utc) {
    Write-Host "FAIL: DateTimeKind mismatch, expected Utc, got: $($utc.Kind)"
    exit 1
}

Write-Host "PASS"
exit 0
