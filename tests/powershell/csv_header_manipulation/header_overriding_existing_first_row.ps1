# vybe-test: powershell/csv_header_manipulation/header_overriding_existing_first_row
$csv = @"
OldCol1,OldCol2
Val1,Val2
"@
$rows = @($csv | ConvertFrom-Csv -Header "NewA", "NewB")
# When -Header is supplied, first row is treated as data
if ($rows.Length -ne 2 -or $rows[0].NewA -ne "OldCol1" -or $rows[1].NewA -ne "Val1") {
    Write-Host "FAIL: Header overriding existing first row failed, got count $($rows.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
