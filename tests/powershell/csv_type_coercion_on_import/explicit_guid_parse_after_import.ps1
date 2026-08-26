# vybe-test: powershell/csv_type_coercion_on_import/explicit_guid_parse_after_import
$gStr = "12345678-1234-1234-1234-123456789abc"
$csv = "Id,Name`n$gStr,Target"
$row = $csv | ConvertFrom-Csv
$g = [guid]::Parse($row.Id)
if ($g -ne [guid]::Parse($gStr)) {
    Write-Host "FAIL: Explicit GUID parse after import failed"
    exit 1
}
Write-Host "PASS"
exit 0
