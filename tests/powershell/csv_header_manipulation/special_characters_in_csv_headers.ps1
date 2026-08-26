# vybe-test: powershell/csv_header_manipulation/special_characters_in_csv_headers
$csv = "ColA,ColB`n100,Prod"
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0].ColA -ne "100" -or $rows[0].ColB -ne "Prod") {
    Write-Host "FAIL: CSV header parsing failed"
    exit 1
}
Write-Host "PASS"
exit 0
