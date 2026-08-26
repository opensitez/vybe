# vybe-test: powershell/type_unsigned_integers/explicit_cast_to_uint64
$val = [uint64]"9876543210123"
if ($val.GetType().Name -ne "UInt64" -or $val -ne 9876543210123) {
    Write-Host "FAIL: [uint64] cast failed"
    exit 1
}
Write-Host "PASS"
exit 0
