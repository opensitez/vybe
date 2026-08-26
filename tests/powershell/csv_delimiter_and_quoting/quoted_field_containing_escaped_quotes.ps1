# vybe-test: powershell/csv_delimiter_and_quoting/quoted_field_containing_escaped_quotes
$csv = @"
Code,Comment
100,"He said ""Hello"" to everyone"
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0].Comment -ne 'He said "Hello" to everyone') {
    Write-Host "FAIL: Quoted field containing escaped quotes failed, got '$($rows[0].Comment)'"
    exit 1
}
Write-Host "PASS"
exit 0
