# vybe-test: powershell/type_accelerators/type_accelerator_bigint
$big = [bigint]"123456789012345678901234567890"
$doubled = $big * 2
if ($doubled.ToString() -ne "246913578024691357802469135780") {
    Write-Host "FAIL: BigInt multiplication result incorrect, got $($doubled.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
