# vybe-test: powershell/csv_delimiter_and_quoting/custom_pipe_delimiter_import
$csv = @"
Id|City|Country
1|Paris|France
2|Tokyo|Japan
"@
$rows = @($csv | ConvertFrom-Csv -Delimiter '|')
if ($rows.Length -ne 2 -or $rows[0].City -ne "Paris" -or $rows[1].Country -ne "Japan") {
    Write-Host "FAIL: Pipe delimiter import failed"
    exit 1
}
Write-Host "PASS"
exit 0
