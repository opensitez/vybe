# vybe-test: powershell/type_unsigned_integers/parse_uint64_from_decimal_string
$val = [uint64]::Parse("12345678901234567890")
if ($val.ToString() -ne "12345678901234567890") {
    Write-Host "FAIL: uint64 Parse mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
