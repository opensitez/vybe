# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_displayhint_date_noteproperty
# -DisplayHint Date attaches a DisplayHint NoteProperty with value 'Date' to the DateTime instance
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$hintObj = Get-Date -Date $date -DisplayHint Date

if ($hintObj.DisplayHint -ne "Date") {
    Write-Host "FAIL: expected DisplayHint property 'Date', got: '$($hintObj.DisplayHint)'"
    exit 1
}

Write-Host "PASS"
exit 0
