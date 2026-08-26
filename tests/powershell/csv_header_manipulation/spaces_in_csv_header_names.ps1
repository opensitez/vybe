# vybe-test: powershell/csv_header_manipulation/spaces_in_csv_header_names
$csv = @"
"First Name","Last Name"
Alice,Smith
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0]."First Name" -ne "Alice" -or $rows[0]."Last Name" -ne "Smith") {
    Write-Host "FAIL: Spaces in CSV header names failed"
    exit 1
}
Write-Host "PASS"
exit 0
