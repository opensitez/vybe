# vybe-test: powershell/test_path_cmdlet/test_path_older_than_future_date_true
# -OlderThan returns $true when tested against a date in the future
$futureDate = (Get-Date).AddYears(1)
$res = Test-Path "Cargo.toml" -OlderThan $futureDate

if ($res -ne $true) {
    Write-Host "FAIL: file expected to be older than future date, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
