# vybe-test: powershell/type_unsigned_integers/explicit_cast_to_uint32
$val = [uint32]"123456"
if ($val.GetType().Name -ne "UInt32" -or $val -ne 123456) {
    Write-Host "FAIL: [uint32] cast failed"
    exit 1
}
Write-Host "PASS"
exit 0
