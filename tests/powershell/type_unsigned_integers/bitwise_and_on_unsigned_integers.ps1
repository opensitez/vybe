# vybe-test: powershell/type_unsigned_integers/bitwise_and_on_unsigned_integers
[uint32]$a = 0xF0F0F0F0
[uint32]$b = 0x0F0F0F0F
$res = $a -band $b
if ($res -ne 0) {
    Write-Host "FAIL: bitwise AND on disjoint masks must be 0"
    exit 1
}
Write-Host "PASS"
exit 0
