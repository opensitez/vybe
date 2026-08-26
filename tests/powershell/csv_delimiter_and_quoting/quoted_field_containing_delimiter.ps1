# vybe-test: powershell/csv_delimiter_and_quoting/quoted_field_containing_delimiter
$csv = @"
Name,Description
Widget,"A small, fast widget"
Gadget,"A large, heavy gadget"
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows.Length -ne 2 -or $rows[0].Description -ne "A small, fast widget") {
    Write-Host "FAIL: Quoted field containing delimiter failed, got '$($rows[0].Description)'"
    exit 1
}
Write-Host "PASS"
exit 0
