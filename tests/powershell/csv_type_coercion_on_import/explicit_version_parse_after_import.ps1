# vybe-test: powershell/csv_type_coercion_on_import/explicit_version_parse_after_import
$csv = @"
Module,Version
Vybe,1.2.3
"@
$row = $csv | ConvertFrom-Csv
$ver = [version]::Parse($row.Version)
if ($ver.Major -ne 1 -or $ver.Minor -ne 2 -or $ver.Build -ne 3) {
    Write-Host "FAIL: Explicit Version parse after import failed"
    exit 1
}
Write-Host "PASS"
exit 0
