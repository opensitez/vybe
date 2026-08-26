# vybe-test: powershell/csv_header_manipulation/unicode_characters_in_headers
$csv = @"
"Caf`u{00E9}","Na`u{00EF}ve"
Paris,Yes
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0]."Caf`u{00E9}" -ne "Paris") {
    Write-Host "FAIL: Unicode characters in headers failed"
    exit 1
}
Write-Host "PASS"
exit 0
