# vybe-test: powershell/test_path_cmdlet/test_path_newer_than_future_date_false
# -NewerThan returns $false when tested against a date in the distant future
$futureDate = (Get-Date).AddYears(1)
$res = Test-Path "Cargo.toml" -NewerThan $futureDate

if ($res -ne $false) {
    Write-Host "FAIL: file expected to NOT be newer than future date, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
