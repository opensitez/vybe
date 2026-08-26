# vybe-test: powershell/parameters_validate_script/validatescript_date_in_the_past
function Set-PastDate {
    param([ValidateScript({ $_ -lt [datetime]::UtcNow })][datetime]$Date)
    return $Date.Year
}
$dt = [datetime]::Parse("2020-01-01")
$res = Set-PastDate -Date $dt
if ($res -ne 2020) {
    Write-Host "FAIL: ValidateScript past date check failed"
    exit 1
}
Write-Host "PASS"
exit 0
