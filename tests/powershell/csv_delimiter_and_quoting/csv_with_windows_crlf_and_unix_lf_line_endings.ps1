# vybe-test: powershell/csv_delimiter_and_quoting/csv_with_windows_crlf_and_unix_lf_line_endings
$csv = "A,B`r`n1,2`n3,4`r`n"
$rows = @($csv | ConvertFrom-Csv)
if ($rows.Length -ne 2 -or $rows[1].A -ne "3") {
    Write-Host "FAIL: Mixed line endings CSV failed"
    exit 1
}
Write-Host "PASS"
exit 0
