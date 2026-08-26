# vybe-test: powershell/csv_delimiter_and_quoting/quoted_field_containing_multiline_newlines
$csv = @"
Id,Notes
1,"Line 1
Line 2"
"@
$rows = @($csv | ConvertFrom-Csv)
if (-not $rows[0].Notes.Contains("`n") -or -not $rows[0].Notes.Contains("Line 2")) {
    Write-Host "FAIL: Multiline quoted field failed, got '$($rows[0].Notes)'"
    exit 1
}
Write-Host "PASS"
exit 0
