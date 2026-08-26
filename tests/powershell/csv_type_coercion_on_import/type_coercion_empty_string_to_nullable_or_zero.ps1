# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_empty_string_to_nullable_or_zero
$csv = @"
Val
""
"@
$row = $csv | ConvertFrom-Csv
$parsedInt = 0
$ok = [int]::TryParse($row.Val, [ref]$parsedInt)
if ($ok -ne $false) {
    Write-Host "FAIL: TryParse on empty string should return false"
    exit 1
}
Write-Host "PASS"
exit 0
