# vybe-test: powershell/csv_header_manipulation/empty_header_name_in_custom_headers_handled
$csv = "1,2`n3,4"
$rows = @($csv | ConvertFrom-Csv -Header "A", "B")
if ($rows.Length -ne 2 -or $rows[0].A -ne "1") {
    Write-Host "FAIL: Header handling failed"
    exit 1
}
Write-Host "PASS"
exit 0
