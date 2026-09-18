# vybe-test: powershell/culture_and_globalization_cmdlets/culture_specific_date_parsing_with_datetimeformat
# Culture-specific DateTimeFormat enables parsing of locale-formatted date strings
$de = Get-Culture -Name "de-DE"

# German date format: dd.MM.yyyy
$parsed = [DateTime]::Parse("10.05.2026", $de.DateTimeFormat)

if ($parsed.Year -ne 2026 -or $parsed.Month -ne 5 -or $parsed.Day -ne 10) {
    Write-Host "FAIL: German date parse mismatch, got $($parsed.ToString('yyyy-MM-dd'))"
    exit 1
}

Write-Host "PASS"
exit 0
