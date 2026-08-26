# vybe-test: powershell/csv_header_manipulation/header_count_greater_than_column_count_assigns_nulls
$csv = @"
1,2
3,4
"@
$rows = @($csv | ConvertFrom-Csv -Header "C1", "C2", "C3")
if ($rows[0].C1 -ne "1" -or $rows[0].C2 -ne "2" -or $rows[0].C3 -ne $null) {
    Write-Host "FAIL: Header count greater than column count failed, C3='$($rows[0].C3)'"
    exit 1
}
Write-Host "PASS"
exit 0
