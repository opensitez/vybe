# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_displayhint_datetime_noteproperty
# -DisplayHint DateTime attaches a DisplayHint NoteProperty with value 'DateTime'
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$hintObj = Get-Date -Date $date -DisplayHint DateTime

if ($hintObj.DisplayHint -ne "DateTime") {
    Write-Host "FAIL: expected DisplayHint property 'DateTime', got: '$($hintObj.DisplayHint)'"
    exit 1
}

Write-Host "PASS"
exit 0
