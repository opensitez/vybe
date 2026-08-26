# vybe-test: powershell/csv_header_manipulation/custom_headers_with_semicolon_delimiter
$csv = @"
1;2;3
4;5;6
"@
$rows = @($csv | ConvertFrom-Csv -Delimiter ';' -Header "A", "B", "C")
if ($rows[0].A -ne "1" -or $rows[1].C -ne "6") {
    Write-Host "FAIL: Custom headers with semicolon delimiter failed"
    exit 1
}
Write-Host "PASS"
exit 0
