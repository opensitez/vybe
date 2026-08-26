# vybe-test: powershell/csv_type_coercion_on_import/explicit_boolean_parse_after_import
$csv = @"
Name,Enabled
Item1,True
Item2,False
"@
$rows = @($csv | ConvertFrom-Csv)
$b1 = [bool]::Parse($rows[0].Enabled)
$b2 = [bool]::Parse($rows[1].Enabled)
if ($b1 -ne $true -or $b2 -ne $false) {
    Write-Host "FAIL: Explicit boolean parse after import failed"
    exit 1
}
Write-Host "PASS"
exit 0
