# vybe-test: powershell/csv_type_coercion_on_import/enum_coercion_after_import
enum UserRole { Guest; Member; Admin }
$csv = @"
Name,Role
Alice,Admin
Bob,Guest
"@
$rows = @($csv | ConvertFrom-Csv)
$r1 = [UserRole]$rows[0].Role
$r2 = [UserRole]$rows[1].Role
if ($r1 -ne [UserRole]::Admin -or $r2 -ne [UserRole]::Guest) {
    Write-Host "FAIL: Enum coercion after import failed"
    exit 1
}
Write-Host "PASS"
exit 0
