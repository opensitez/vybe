# vybe-test: powershell/csv_delimiter_and_quoting/empty_fields_between_delimiters
$csv = "A,B,C`n1,,3`n,2,"
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0].B -ne "" -or $rows[1].B -ne "2") {
    Write-Host "FAIL: Empty fields between delimiters failed"
    exit 1
}
Write-Host "PASS"
exit 0
