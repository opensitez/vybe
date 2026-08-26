# vybe-test: powershell/csv_delimiter_and_quoting/csv_with_trailing_empty_delimiter
$csv = @"
A,B,
1,2,
"@
$rows = @($csv | ConvertFrom-Csv)
# Trailing comma creates an empty-name column or error depending on parser
if ($rows.Length -ne 1) {
    Write-Host "FAIL: Trailing empty delimiter import failed"
    exit 1
}
Write-Host "PASS"
exit 0
