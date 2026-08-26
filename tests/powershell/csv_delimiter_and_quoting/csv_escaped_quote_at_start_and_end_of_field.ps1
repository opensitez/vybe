# vybe-test: powershell/csv_delimiter_and_quoting/csv_escaped_quote_at_start_and_end_of_field
$csv = @"
Text
"""Start and End"""
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0].Text -ne '"Start and End"') {
    Write-Host "FAIL: Quotes at start and end of field failed, got '$($rows[0].Text)'"
    exit 1
}
Write-Host "PASS"
exit 0
