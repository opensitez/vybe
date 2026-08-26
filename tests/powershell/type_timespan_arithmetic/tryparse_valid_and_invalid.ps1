# vybe-test: powershell/type_timespan_arithmetic/tryparse_valid_and_invalid
$outTs = [timespan]::Zero
$valid = [timespan]::TryParse("01:23:45", [ref]$outTs)
$invalid = [timespan]::TryParse("not-a-timespan", [ref]$outTs)
if (-not $valid -or $invalid) {
    Write-Host "FAIL: TryParse validation check"
    exit 1
}
Write-Host "PASS"
exit 0
