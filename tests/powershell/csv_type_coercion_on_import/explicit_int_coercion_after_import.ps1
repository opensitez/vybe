# vybe-test: powershell/csv_type_coercion_on_import/explicit_int_coercion_after_import
$csv = @"
Id,Qty
10,50
"@
$row = $csv | ConvertFrom-Csv
$id = [int]$row.Id
$qty = [int]$row.Qty
$total = $id * $qty
if ($total -ne 500) {
    Write-Host "FAIL: Explicit int coercion after import failed, got $total"
    exit 1
}
Write-Host "PASS"
exit 0
