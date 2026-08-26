# vybe-test: powershell/json_date_time_formats/date_formatted_with_culture_invariant
$dt = [datetime]::Parse("2026-08-26 12:00:00", [System.Globalization.CultureInfo]::InvariantCulture)
$json = @{ InvDate = $dt.ToString("O", [System.Globalization.CultureInfo]::InvariantCulture) } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
if (-not $obj.InvDate.Contains("2026-08-26")) {
    Write-Host "FAIL: Culture invariant date in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
