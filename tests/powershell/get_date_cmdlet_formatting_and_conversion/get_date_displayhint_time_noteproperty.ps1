# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_displayhint_time_noteproperty
# -DisplayHint Time attaches a DisplayHint NoteProperty with value 'Time' to the DateTime instance
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$hintObj = Get-Date -Date $date -DisplayHint Time

if ($hintObj.DisplayHint -ne "Time") {
    Write-Host "FAIL: expected DisplayHint property 'Time', got: '$($hintObj.DisplayHint)'"
    exit 1
}

Write-Host "PASS"
exit 0
