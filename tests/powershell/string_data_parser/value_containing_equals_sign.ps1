# vybe-test: powershell/string_data_parser/value_containing_equals_sign
$str = "connectionString = server=db;database=app;uid=user;"
$ht = ConvertFrom-StringData -StringData $str
if ($ht["connectionString"] -ne "server=db;database=app;uid=user;") {
    Write-Host "FAIL: Value containing equals sign failed, got '$($ht["connectionString"])'"
    exit 1
}
Write-Host "PASS"
exit 0
