# vybe-test: powershell/culture_and_globalization_cmdlets/culture_datetimeformat_german_short_date_pattern
# The German (de-DE) culture defines 'dd.MM.yyyy' as its ShortDatePattern
$de = Get-Culture -Name "de-DE"
$pattern = $de.DateTimeFormat.ShortDatePattern

if ($pattern -ne "dd.MM.yyyy") {
    Write-Host "FAIL: expected German short date 'dd.MM.yyyy', got: '$pattern'"
    exit 1
}

Write-Host "PASS"
exit 0
