# vybe-test: powershell/csv_delimiter_and_quoting/single_column_csv
$csv = @"
Item
Apple
Banana
Cherry
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows.Length -ne 3 -or $rows[1].Item -ne "Banana") {
    Write-Host "FAIL: Single column CSV import failed"
    exit 1
}
Write-Host "PASS"
exit 0
