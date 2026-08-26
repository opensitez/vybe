# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_array_split_from_delimited_field
$csv = @"
User,Roles
Alice,Admin;User;Tester
"@
$row = $csv | ConvertFrom-Csv
$roles = $row.Roles.Split(';')
if ($roles.Length -ne 3 -or $roles[0] -ne "Admin" -or $roles[2] -ne "Tester") {
    Write-Host "FAIL: Delimited field array split failed"
    exit 1
}
Write-Host "PASS"
exit 0
